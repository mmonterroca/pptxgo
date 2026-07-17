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
	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// TextBox is a handle onto a text-box shape's txBody, returned by
// Slide.AddTextBox. Its only child method is AddParagraph — text and
// formatting are set per paragraph.
type TextBox struct {
	pres *Presentation
	body *drawingml.TextBody
}

// AddParagraph appends a new, empty paragraph to the text box and returns a
// handle for adding runs and formatting to it.
func (tb *TextBox) AddParagraph() *Paragraph {
	p := &drawingml.Paragraph{}
	tb.body.Paragraphs = append(tb.body.Paragraphs, p)
	return &Paragraph{pres: tb.pres, p: p}
}

// Paragraph is a handle onto a single a:p, returned by TextBox.AddParagraph.
// Text starts a new run; the run-formatting methods (Bold, Italic,
// Underline, FontSize, Font, Color) apply to the most recently started run,
// so a single paragraph can mix differently-formatted runs by calling Text
// again in between. Alignment applies to the paragraph as a whole and can
// be called at any point in the chain.
type Paragraph struct {
	pres   *Presentation
	p      *drawingml.Paragraph
	curRun *drawingml.Run
}

// Text starts a new run of text within the paragraph. Subsequent
// formatting calls (Bold, Italic, ...) apply to this run until Text is
// called again.
func (pg *Paragraph) Text(s string) *Paragraph {
	run := &drawingml.Run{Text: drawingml.NewText(s)}
	pg.p.Runs = append(pg.p.Runs, run)
	pg.curRun = run
	return pg
}

// rPr returns the run-properties struct for the current run, allocating it
// on first use. It returns nil if no run has been started yet (Bold and
// friends called before Text), in which case the formatting call is a
// no-op rather than a panic.
func (pg *Paragraph) rPr() *drawingml.RPr {
	if pg.curRun == nil {
		return nil
	}
	if pg.curRun.RPr == nil {
		pg.curRun.RPr = &drawingml.RPr{}
	}
	return pg.curRun.RPr
}

// Bold makes the current run bold.
func (pg *Paragraph) Bold() *Paragraph {
	if rpr := pg.rPr(); rpr != nil {
		rpr.B = true
	}
	return pg
}

// Italic makes the current run italic.
func (pg *Paragraph) Italic() *Paragraph {
	if rpr := pg.rPr(); rpr != nil {
		rpr.I = true
	}
	return pg
}

// Underline underlines the current run with a single line.
func (pg *Paragraph) Underline() *Paragraph {
	if rpr := pg.rPr(); rpr != nil {
		rpr.U = "sng"
	}
	return pg
}

// FontSize sets the current run's font size, in points (1-4000, matching
// the schema's centipoint range of 100-400000). An out-of-range value is
// recorded as an error on the presentation (returned by Save) and leaves
// the run's size unset.
func (pg *Paragraph) FontSize(points int) *Paragraph {
	if points < 1 || points > 4000 {
		pg.pres.addErr(errors.InvalidArgument("Paragraph.FontSize", "points", points, "must be between 1 and 4000"))
		return pg
	}
	if rpr := pg.rPr(); rpr != nil {
		rpr.Sz = points * 100
	}
	return pg
}

// Font sets the current run's Latin-script typeface.
func (pg *Paragraph) Font(name string) *Paragraph {
	if rpr := pg.rPr(); rpr != nil {
		rpr.Latin = &drawingml.Latin{Typeface: name}
	}
	return pg
}

// Color sets the current run's text color to a solid fill.
func (pg *Paragraph) Color(c drawingml.Color) *Paragraph {
	if rpr := pg.rPr(); rpr != nil {
		rpr.SolidFill = drawingml.NewSolidFillRGB(c)
	}
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
