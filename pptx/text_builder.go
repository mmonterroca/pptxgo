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

// TextBox is a handle onto a text-box shape, returned by Slide.AddTextBox.
// AddParagraph adds text content; Fill and Border set the shape's own
// background and outline.
type TextBox struct {
	pres *Presentation
	body *drawingml.TextBody
	spPr *SpPr
}

// AddParagraph appends a new, empty paragraph to the text box and returns a
// handle for adding runs and formatting to it.
func (tb *TextBox) AddParagraph() *Paragraph {
	p := &drawingml.Paragraph{}
	tb.body.Paragraphs = append(tb.body.Paragraphs, p)
	return &Paragraph{pres: tb.pres, p: p}
}

// Fill sets the text box's background to a solid color.
func (tb *TextBox) Fill(c drawingml.Color) *TextBox {
	tb.spPr.Fill = drawingml.NewSolidFillRGB(c)
	return tb
}

// Border sets the text box's outline to a solid color at the given width,
// in points (e.g. 0.75, 1.5; 0-1584, matching ST_LineWidth's 0-20,116,800
// EMU range). An out-of-range width is recorded as an error on the
// presentation (returned by Save) and leaves the border unset.
func (tb *TextBox) Border(c drawingml.Color, widthPoints float64) *TextBox {
	tb.spPr.Ln = newLn(tb.pres, c, widthPoints)
	return tb
}

// maxLineWidthPoints is ST_LineWidth's maximum, 20,116,800 EMU, expressed
// in points (20116800 / EMUsPerPoint).
const maxLineWidthPoints = 1584

// newLn builds a solid-color a:ln of the given width in points, shared by
// TextBox.Border and PictureRef.Border. Point-to-EMU conversion happens
// here, at the fluent-API boundary, exactly once — the same pattern
// Paragraph.FontSize uses for centipoints. An out-of-range width (negative,
// or over ST_LineWidth's maximum) is recorded as an error on pres and
// returns nil, leaving the caller's Ln field unset rather than emitting a
// schema-invalid a:ln/@w.
func newLn(pres *Presentation, c drawingml.Color, widthPoints float64) *drawingml.Ln {
	if widthPoints < 0 || widthPoints > maxLineWidthPoints {
		pres.addErr(errors.InvalidArgument("Border", "widthPoints", widthPoints,
			"must be between 0 and 1584 (ST_LineWidth's 0-20,116,800 EMU range)"))
		return nil
	}
	return drawingml.NewLn(c, int(widthPoints*drawingml.EMUsPerPoint))
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
	pres    *Presentation
	p       *drawingml.Paragraph
	curRuns []*drawingml.Run
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

// Alignment sets the paragraph's horizontal text alignment.
func (pg *Paragraph) Alignment(a Alignment) *Paragraph {
	if pg.p.PPr == nil {
		pg.p.PPr = &drawingml.PPr{}
	}
	pg.p.PPr.Algn = string(a)
	return pg
}
