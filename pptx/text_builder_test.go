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
	"bytes"
	"strings"
	"testing"
)

func TestFontSize_OutOfRangeIsAccumulatedNotPanicked(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))

	// The chain must keep returning usable handles even after an invalid
	// value — the deferred-error pattern only pays off if a long chain
	// doesn't need an `if err != nil` after every call.
	tb.AddParagraph().Text("too big").FontSize(5000).Bold()

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated FontSize error")
	}
}

func TestFontSize_ValidRangeDoesNotError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("fine").FontSize(1).FontSize(4000)

	if err := p.Save(&bytes.Buffer{}); err != nil {
		t.Fatalf("expected no error for boundary-valid font sizes, got %v", err)
	}
}

func TestParagraph_FormattingWithNoTextIsANoOp(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))

	// Bold/FontSize/etc. called before Text must not panic — there is no
	// current run to attach formatting to yet.
	tb.AddParagraph().Bold().FontSize(32).Alignment(AlignCenter)

	if err := p.Save(&bytes.Buffer{}); err != nil {
		t.Fatalf("expected no error, got %v", err)
	}
}

func TestFontSize_OutOfRangeBeforeTextIsANoOpNotAnError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))

	// FontSize(5000) is out of range, but it's called before Text — every
	// other formatting method documents this as a no-op, and FontSize's own
	// range check must not run ahead of that "no current run" guard.
	tb.AddParagraph().FontSize(5000).Text("ok")

	if err := p.Save(&bytes.Buffer{}); err != nil {
		t.Fatalf("expected no error (FontSize before Text is a no-op), got %v", err)
	}
}

func TestFontSize_AcceptsHalfPoints(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("half point").FontSize(10.5)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])
	if !strings.Contains(slide, `sz="1050"`) {
		t.Errorf("expected sz=\"1050\" (10.5pt) in slide1.xml, got %s", slide)
	}
}

func TestText_NewlineInsertsExplicitLineBreak(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("Line1\nLine2").Bold()

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	line1Idx := strings.Index(slide, "Line1")
	brIdx := strings.Index(slide, "<a:br")
	line2Idx := strings.Index(slide, "Line2")
	if line1Idx == -1 || brIdx == -1 || line2Idx == -1 {
		t.Fatalf("expected Line1, a:br, and Line2 all present, got %s", slide)
	}
	if !(line1Idx < brIdx && brIdx < line2Idx) {
		t.Errorf("expected Line1 < a:br < Line2 order, got %s", slide)
	}
	// Bold, called once after both lines were started by the same Text
	// call, must apply to both runs, not just the last.
	if strings.Count(slide, `b="1"`) != 2 {
		t.Errorf("expected both lines to be bold, got %s", slide)
	}
}
