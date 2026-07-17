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
