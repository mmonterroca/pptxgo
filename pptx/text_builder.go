/*
MIT License

Copyright (c) 2026 Misael Monterroca

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

package pptx

import (
	"math"
	"strings"

	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// ShapeRef is a handle onto any p:sp shape on a slide — a text box
// (Slide.AddTextBox) or an arbitrary autoshape (Slide.AddShape). AddParagraph
// adds text content; Fill and Border set the shape's own background and
// outline; Rotation, FlipH, and FlipV set its transform. TextBox is an alias:
// a text box is simply a ShapeRef whose geometry defaults to rect.
type ShapeRef struct {
	pres      *Presentation
	slidePath string // the owning slide's part path, needed for per-slide hyperlink relationships (see Paragraph.Hyperlink)
	body      *drawingml.TextBody
	spPr      *SpPr
}

// TextBox is a handle onto a text-box shape, returned by Slide.AddTextBox.
type TextBox = ShapeRef

// AddParagraph appends a new, empty paragraph to the shape and returns a
// handle for adding runs and formatting to it.
func (sr *ShapeRef) AddParagraph() *Paragraph {
	p := &drawingml.Paragraph{}
	sr.body.Paragraphs = append(sr.body.Paragraphs, p)
	return &Paragraph{pres: sr.pres, slidePath: sr.slidePath, p: p}
}

// Fill sets the shape's background to a solid color.
func (sr *ShapeRef) Fill(c drawingml.Color) *ShapeRef {
	sr.spPr.NoFill = nil
	sr.spPr.Fill = drawingml.NewSolidFillRGB(c)
	return sr
}

// FillScheme sets the shape's background to a theme color, referenced by
// scheme slot (e.g. SchemeAccent1) rather than an explicit RGB value.
func (sr *ShapeRef) FillScheme(scheme SchemeColor) *ShapeRef {
	sr.spPr.NoFill = nil
	sr.spPr.Fill = drawingml.NewSolidFillScheme(string(scheme))
	return sr
}

// NoFill removes the shape's fill entirely — distinct from never calling
// Fill, which lets the shape inherit one from its style or layout instead.
func (sr *ShapeRef) NoFill() *ShapeRef {
	sr.spPr.Fill = nil
	sr.spPr.NoFill = &drawingml.NoFill{}
	return sr
}

// Border sets the shape's outline to a solid color at the given width,
// in points (e.g. 0.75, 1.5; 0-1584, matching ST_LineWidth's 0-20,116,800
// EMU range). An out-of-range width is recorded as an error on the
// presentation (returned by Save) and leaves the border unset.
func (sr *ShapeRef) Border(c drawingml.Color, widthPoints float64) *ShapeRef {
	sr.spPr.Ln = newLn(sr.pres, c, widthPoints)
	return sr
}

// BorderScheme sets the shape's outline to a theme color, referenced by
// scheme slot (e.g. SchemeAccent1), at the given width in points — see
// Border for the width's valid range.
func (sr *ShapeRef) BorderScheme(scheme SchemeColor, widthPoints float64) *ShapeRef {
	sr.spPr.Ln = newLnScheme(sr.pres, scheme, widthPoints)
	return sr
}

// Rotation sets the shape's rotation, in degrees clockwise (e.g. 45, -90;
// any value works, including beyond a full turn — 405 is the same
// rotation as 45). AddShape and AddTextBox always give a shape its own
// a:xfrm, so this is never a no-op for them; a placeholder from
// Slide.AddPlaceholder/Title/Body has no a:xfrm of its own (see WithLayout)
// and so has nothing to rotate — see the shared nil-Xfrm handling this
// calls into.
func (sr *ShapeRef) Rotation(degrees float64) *ShapeRef {
	if math.IsNaN(degrees) || math.IsInf(degrees, 0) {
		sr.pres.addErr(errors.InvalidArgument("Rotation", "degrees", degrees, "must be a finite number"))
		return sr
	}
	if !sr.requireXfrm("Rotation") {
		return sr
	}
	// Normalize to (-360, 360) before converting to 60,000ths of a degree
	// (a:xfrm/@rot): rotation is inherently modular (405 degrees looks
	// identical to 45), so this changes nothing visually, but an
	// un-normalized large input — Rotation(36000), say — would multiply
	// out to a value ST_Angle's underlying xsd:int (32-bit signed) can't
	// hold, corrupting the file silently since Save has no other way to
	// detect it.
	sr.spPr.Xfrm.Rot = int(math.Round(math.Mod(degrees, 360) * 60000))
	return sr
}

// FlipH flips the shape horizontally.
func (sr *ShapeRef) FlipH() *ShapeRef {
	if !sr.requireXfrm("FlipH") {
		return sr
	}
	sr.spPr.Xfrm.FlipH = true
	return sr
}

// FlipV flips the shape vertically.
func (sr *ShapeRef) FlipV() *ShapeRef {
	if !sr.requireXfrm("FlipV") {
		return sr
	}
	sr.spPr.Xfrm.FlipV = true
	return sr
}

// requireXfrm reports whether sr already has an a:xfrm to set a transform
// attribute on. A nil Xfrm means the shape is a placeholder that
// deliberately omits its own a:xfrm to inherit position/size from its
// layout (see Slide.AddPlaceholder) — silently no-oping a transform call
// on it would discard the caller's request without a trace, so this
// records an error on the presentation (returned by Save) instead.
func (sr *ShapeRef) requireXfrm(method string) bool {
	if sr.spPr.Xfrm != nil {
		return true
	}
	sr.pres.addErr(errors.InvalidArgument(method, "shape", "placeholder",
		"a placeholder with inherited geometry (see Slide.AddPlaceholder) has no a:xfrm to set a transform on"))
	return false
}

// WordWrap sets whether text wraps at the shape's edge. PowerPoint's own
// default is wrapping enabled, so this only needs calling to disable it.
func (sr *ShapeRef) WordWrap(enabled bool) *ShapeRef {
	if enabled {
		sr.body.BodyPr.Wrap = "square"
	} else {
		sr.body.BodyPr.Wrap = "none"
	}
	return sr
}

// Anchor sets the text body's vertical alignment within the shape.
func (sr *ShapeRef) Anchor(a VerticalAnchor) *ShapeRef {
	sr.body.BodyPr.Anchor = string(a)
	return sr
}

// Insets sets the text body's internal margins (the gap between the
// shape's outline and its text), all in points.
func (sr *ShapeRef) Insets(left, top, right, bottom float64) *ShapeRef {
	sr.body.BodyPr.LIns = emuPtr(left)
	sr.body.BodyPr.TIns = emuPtr(top)
	sr.body.BodyPr.RIns = emuPtr(right)
	sr.body.BodyPr.BIns = emuPtr(bottom)
	return sr
}

// emuPtr converts points to EMUs, rounding to the nearest whole EMU, and
// returns a pointer to the result. BodyPr's insets and PPr's marL/indent
// are *int specifically so an explicit zero (e.g. Insets(0, 0, 0, 0))
// marshals as e.g. lIns="0" instead of being indistinguishable from
// "never set" and dropped by omitempty.
func emuPtr(points float64) *int {
	v := int(math.Round(points * drawingml.EMUsPerPoint))
	return &v
}

// Autofit sets how the shape's text behaves when it overflows the shape's
// bounds.
func (sr *ShapeRef) Autofit(mode AutofitMode) *ShapeRef {
	sr.body.BodyPr.NoAutofit, sr.body.BodyPr.NormAutofit, sr.body.BodyPr.SpAutoFit = nil, nil, nil
	switch mode {
	case AutofitNone:
		sr.body.BodyPr.NoAutofit = &drawingml.NoAutofit{}
	case AutofitShrinkText:
		sr.body.BodyPr.NormAutofit = &drawingml.NormAutofit{}
	case AutofitResizeShape:
		sr.body.BodyPr.SpAutoFit = &drawingml.SpAutoFit{}
	}
	return sr
}

// maxLineWidthPoints is ST_LineWidth's maximum, 20,116,800 EMU, expressed
// in points (20116800 / EMUsPerPoint).
const maxLineWidthPoints = 1584

// newLn builds a solid-color a:ln of the given width in points, shared by
// ShapeRef.Border and PictureRef.Border. Point-to-EMU conversion happens
// here, at the fluent-API boundary, exactly once — the same pattern
// Paragraph.FontSize uses for centipoints.
func newLn(pres *Presentation, c drawingml.Color, widthPoints float64) *drawingml.Ln {
	w, ok := validatedLineWidthEMU(pres, widthPoints)
	if !ok {
		return nil
	}
	return drawingml.NewLn(c, w)
}

// newLnScheme is newLn's theme-color counterpart, shared by
// ShapeRef.BorderScheme.
func newLnScheme(pres *Presentation, scheme SchemeColor, widthPoints float64) *drawingml.Ln {
	w, ok := validatedLineWidthEMU(pres, widthPoints)
	if !ok {
		return nil
	}
	return drawingml.NewLnScheme(string(scheme), w)
}

// validatedLineWidthEMU converts widthPoints to EMUs (rounding to the
// nearest whole EMU, not truncating — the same class of float64 precision
// issue FontSize's centipoint conversion already guards against), or
// records an out-of-range width as an error on pres and returns ok=false,
// leaving the caller's Ln field unset rather than emitting a
// schema-invalid a:ln/@w. NaN fails both bounds comparisons below (every
// comparison against NaN is false), so it is checked explicitly first —
// without that check it would silently reach the int() conversion, which
// produces an implementation-defined result for a non-finite float64.
func validatedLineWidthEMU(pres *Presentation, widthPoints float64) (emu int, ok bool) {
	if math.IsNaN(widthPoints) || widthPoints < 0 || widthPoints > maxLineWidthPoints {
		pres.addErr(errors.InvalidArgument("Border", "widthPoints", widthPoints,
			"must be between 0 and 1584 (ST_LineWidth's 0-20,116,800 EMU range)"))
		return 0, false
	}
	return int(math.Round(widthPoints * drawingml.EMUsPerPoint)), true
}

// Paragraph is a handle onto a single a:p, returned by TextBox.AddParagraph.
// Text starts one or more new runs (a "\n" in the given string splits it
// into multiple runs with an explicit a:br between them — PowerPoint does
// not treat a literal newline inside a run's text as a line break). The
// run-formatting methods (Bold, Italic, Underline, FontSize, Font, Color)
// apply to every run Text most recently started, so a single paragraph can
// mix differently-formatted runs by calling Text again in between.
// Alignment applies to the paragraph as a whole and can be called at any
// point in the chain.
type Paragraph struct {
	pres      *Presentation
	slidePath string // the owning slide's part path, needed by Hyperlink to scope its relationship correctly
	p         *drawingml.Paragraph
	curRuns   []*drawingml.Run
}

// Text starts one or more new runs of text within the paragraph, splitting
// on "\n" and inserting an a:br between the resulting lines. Subsequent
// formatting calls (Bold, Italic, ...) apply to every run just started,
// until Text is called again.
func (pg *Paragraph) Text(s string) *Paragraph {
	lines := strings.Split(s, "\n")
	runs := make([]*drawingml.Run, 0, len(lines))
	for i, line := range lines {
		if i > 0 {
			pg.p.Content = append(pg.p.Content, &drawingml.Br{})
		}
		run := &drawingml.Run{Text: drawingml.NewText(line)}
		pg.p.Content = append(pg.p.Content, run)
		runs = append(runs, run)
	}
	pg.curRuns = runs
	return pg
}

// eachRPr calls fn with the run-properties struct of every run Text most
// recently started, allocating each on first use. It calls fn zero times
// if no run has been started yet (Bold and friends called before Text),
// in which case the formatting call is a no-op rather than a panic.
func (pg *Paragraph) eachRPr(fn func(*drawingml.RPr)) {
	for _, run := range pg.curRuns {
		if run.RPr == nil {
			run.RPr = &drawingml.RPr{}
		}
		fn(run.RPr)
	}
}

// Bold makes the current run(s) bold.
func (pg *Paragraph) Bold() *Paragraph {
	pg.eachRPr(func(rpr *drawingml.RPr) { rpr.B = true })
	return pg
}

// Italic makes the current run(s) italic.
func (pg *Paragraph) Italic() *Paragraph {
	pg.eachRPr(func(rpr *drawingml.RPr) { rpr.I = true })
	return pg
}

// Underline underlines the current run(s) with a single line.
func (pg *Paragraph) Underline() *Paragraph {
	pg.eachRPr(func(rpr *drawingml.RPr) { rpr.U = "sng" })
	return pg
}

// FontSize sets the current run(s)' font size, in points (1-4000, matching
// the schema's centipoint range of 100-400000; half-points such as 10.5
// are valid). Formatting calls made before any Text call are a documented
// no-op — including this validation, so FontSize(5000) before Text does
// not record an error. An out-of-range value called with a current run
// present is recorded as an error on the presentation (returned by Save)
// and leaves the run(s)' size unset.
func (pg *Paragraph) FontSize(points float64) *Paragraph {
	if len(pg.curRuns) == 0 {
		return pg
	}
	if points < 1 || points > 4000 {
		pg.pres.addErr(errors.InvalidArgument("Paragraph.FontSize", "points", points, "must be between 1 and 4000"))
		return pg
	}
	sz := int(math.Round(points * 100))
	pg.eachRPr(func(rpr *drawingml.RPr) { rpr.Sz = sz })
	return pg
}

// Font sets the current run(s)' Latin-script typeface.
func (pg *Paragraph) Font(name string) *Paragraph {
	pg.eachRPr(func(rpr *drawingml.RPr) { rpr.Latin = &drawingml.Latin{Typeface: name} })
	return pg
}

// Color sets the current run(s)' text color to a solid fill.
func (pg *Paragraph) Color(c drawingml.Color) *Paragraph {
	pg.eachRPr(func(rpr *drawingml.RPr) { rpr.SolidFill = drawingml.NewSolidFillRGB(c) })
	return pg
}

// ColorScheme sets the current run(s)' text color to a theme color,
// referenced by scheme slot (e.g. SchemeAccent1) rather than an explicit
// RGB value.
func (pg *Paragraph) ColorScheme(scheme SchemeColor) *Paragraph {
	pg.eachRPr(func(rpr *drawingml.RPr) { rpr.SolidFill = drawingml.NewSolidFillScheme(string(scheme)) })
	return pg
}

// Hyperlink makes the current run(s) a clickable hyperlink to the given
// external URL. Like Text and the other run-formatting methods, it applies
// to whichever run(s) the most recent Text call started, and is a no-op
// if called before any Text call. The relationship is scoped to the
// owning slide's own .rels, the same pattern AddImage already established
// for media.
func (pg *Paragraph) Hyperlink(url string) *Paragraph {
	if len(pg.curRuns) == 0 {
		return pg
	}
	rid, err := pg.pres.pkg.Relationships(pg.slidePath).AddHyperlink(url)
	if err != nil {
		pg.pres.addErr(err)
		return pg
	}
	pg.eachRPr(func(rpr *drawingml.RPr) {
		rpr.HlinkClick = &drawingml.HlinkClick{XmlnsR: drawingml.NamespaceRelationships, RID: rid}
	})
	return pg
}

// Alignment sets the paragraph's horizontal text alignment.
func (pg *Paragraph) Alignment(a Alignment) *Paragraph {
	pg.pPr().Algn = string(a)
	return pg
}

// pPr returns the paragraph's a:pPr, allocating it on first use.
func (pg *Paragraph) pPr() *drawingml.PPr {
	if pg.p.PPr == nil {
		pg.p.PPr = &drawingml.PPr{}
	}
	return pg.p.PPr
}

// Level sets the paragraph's outline level (0-8, PowerPoint's UI levels
// 1-9, matching ST_TextIndentLevelType's range), which controls indent
// and bullet inheritance from the list style. A value outside 0-8 is
// recorded as an error on the presentation (returned by Save) and leaves
// the level unset.
func (pg *Paragraph) Level(lvl int) *Paragraph {
	if lvl < 0 || lvl > 8 {
		pg.pres.addErr(errors.InvalidArgument("Paragraph.Level", "lvl", lvl, "must be between 0 and 8"))
		return pg
	}
	pg.pPr().Lvl = &lvl
	return pg
}

// Indent sets the paragraph's left margin and first-line indent, both in
// points. A negative firstLine produces a hanging indent — the common
// bulleted-text layout where the bullet sits to the left of wrapped text.
func (pg *Paragraph) Indent(marginLeft, firstLine float64) *Paragraph {
	pg.pPr().MarL = emuPtr(marginLeft)
	pg.pPr().Indent = emuPtr(firstLine)
	return pg
}

// clearBullet resets every bullet-group field so exactly one of
// Bullet/NumberedBullet/NoBullet takes effect, matching the schema's
// mutually-exclusive EG_TextBullet choice group.
func (pg *Paragraph) clearBullet() {
	pPr := pg.pPr()
	pPr.BuFont, pPr.BuNone, pPr.BuAutoNum, pPr.BuChar = nil, nil, nil, nil
}

// Bullet sets this paragraph's bullet to an explicit character (e.g. "•"),
// drawn in the given font (e.g. "Arial") — PowerPoint needs that font
// declared alongside the character, or the bullet renders as a
// missing-glyph box. Explicit only: this is a per-paragraph override, not
// the master-inherited bullet a placeholder would otherwise pick up.
func (pg *Paragraph) Bullet(char, font string) *Paragraph {
	pg.clearBullet()
	pg.pPr().BuFont = &drawingml.BuFont{Typeface: font}
	pg.pPr().BuChar = &drawingml.BuChar{Char: char}
	return pg
}

// NumberedBullet sets this paragraph's bullet to an automatically numbered
// scheme (e.g. NumArabicPeriod for "1.", "2.", ...).
func (pg *Paragraph) NumberedBullet(scheme NumberingScheme) *Paragraph {
	pg.clearBullet()
	pg.pPr().BuAutoNum = &drawingml.BuAutoNum{Type: string(scheme)}
	return pg
}

// NoBullet explicitly suppresses any bullet for this paragraph.
func (pg *Paragraph) NoBullet() *Paragraph {
	pg.clearBullet()
	pg.pPr().BuNone = &drawingml.BuNone{}
	return pg
}

// LineSpacing sets this paragraph's line spacing as a percentage of single
// spacing (100 = single, 150 = 1.5x, 200 = double).
func (pg *Paragraph) LineSpacing(percent float64) *Paragraph {
	pg.pPr().LnSpc = &drawingml.TextSpacing{SpcPct: &drawingml.SpcPct{Val: int(math.Round(percent * 1000))}}
	return pg
}

// SpaceBefore sets the space before this paragraph, in points.
func (pg *Paragraph) SpaceBefore(points float64) *Paragraph {
	pg.pPr().SpcBef = &drawingml.TextSpacing{SpcPts: &drawingml.SpcPts{Val: int(math.Round(points * 100))}}
	return pg
}

// SpaceAfter sets the space after this paragraph, in points.
func (pg *Paragraph) SpaceAfter(points float64) *Paragraph {
	pg.pPr().SpcAft = &drawingml.TextSpacing{SpcPts: &drawingml.SpcPts{Val: int(math.Round(points * 100))}}
	return pg
}
