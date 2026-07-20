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

func TestFooter_EmitsCenteredFtrPlaceholder(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.Footer("Confidential — Q3 2026")

	slide := string(generateFrom(t, p)["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<p:ph type="ftr"`) {
		t.Errorf("expected a ftr placeholder, got %s", slide)
	}
	if !strings.Contains(slide, "Confidential — Q3 2026") {
		t.Errorf("expected the footer text, got %s", slide)
	}
	if !strings.Contains(slide, `algn="ctr"`) {
		t.Errorf("expected the footer centered, got %s", slide)
	}
}

func TestDateText_EmitsLeftAlignedDtPlaceholder(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.DateText("July 19, 2026")

	slide := string(generateFrom(t, p)["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<p:ph type="dt"`) {
		t.Errorf("expected a dt placeholder, got %s", slide)
	}
	if !strings.Contains(slide, "July 19, 2026") {
		t.Errorf("expected the literal date text, got %s", slide)
	}
}

func TestSlideNumber_EmitsSldNumPlaceholderWithField(t *testing.T) {
	p := New()
	p.AddSlide() // slide 1
	s := p.AddSlide()
	s.SlideNumber() // slide 2

	slide := string(generateFrom(t, p)["ppt/slides/slide2.xml"])

	if !strings.Contains(slide, `<p:ph type="sldNum"`) {
		t.Errorf("expected a sldNum placeholder, got %s", slide)
	}
	if !strings.Contains(slide, `<a:fld id="`) || !strings.Contains(slide, `type="slidenum"`) {
		t.Errorf("expected an a:fld of type slidenum, got %s", slide)
	}
	// Literal fallback is this slide's position (2).
	if !strings.Contains(slide, "<a:t>2</a:t>") {
		t.Errorf("expected the slide-number field's literal fallback to be 2, got %s", slide)
	}
	if !strings.Contains(slide, `algn="r"`) {
		t.Errorf("expected the slide number right-aligned, got %s", slide)
	}
}

func TestFooterRow_PositionedNearBottomOfCanvas(t *testing.T) {
	p := New() // default 16:9, 6858000 EMU tall
	s := p.AddSlide()
	s.Footer("x")

	slide := string(generateFrom(t, p)["ppt/slides/slide1.xml"])

	// The footer band sits in the bottom ~10% of the canvas: y should be well
	// past the vertical midpoint (3429000). With height 6% and 3% bottom gap,
	// y = 6858000 - 411480 - 205740 = 6240780.
	if !strings.Contains(slide, "<a:off x=") {
		t.Fatalf("expected an a:off on the footer placeholder, got %s", slide)
	}
	if !strings.Contains(slide, `y="6240780"`) {
		t.Errorf("expected the footer y near the bottom (6240780 for a 16:9 canvas), got %s", slide)
	}
}

func TestFooterRow_ScalesWithSlideSize(t *testing.T) {
	// 4:3 is a different height (6858000 too — same height as 16:9!), so use a
	// custom size to prove the geometry tracks the canvas.
	p := New(WithSlideSize(SlideSizeStandard4x3Width, 4000000))
	s := p.AddSlide()
	s.SlideNumber()

	slide := string(generateFrom(t, p)["ppt/slides/slide1.xml"])
	// y = 4000000 - 4000000*6/100 - 4000000*3/100 = 4000000 - 240000 - 120000 = 3640000.
	if !strings.Contains(slide, `y="3640000"`) {
		t.Errorf("expected the footer y to scale with a 4000000-EMU-tall canvas (3640000), got %s", slide)
	}
}

func TestFooter_DuplicateOnSameSlideAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.Footer("first")
	s.Footer("second")

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated duplicate-footer error")
	}
}

func TestFooterRow_AllThreeCoexistOnOneSlide(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.DateText("2026-07-19").Footer("Acme Confidential").SlideNumber()

	if err := p.Save(&bytes.Buffer{}); err != nil {
		t.Fatalf("expected date+footer+slidenumber to coexist, got %v", err)
	}
	slide := string(generateFrom(t, p)["ppt/slides/slide1.xml"])
	for _, want := range []string{`type="dt"`, `type="ftr"`, `type="sldNum"`} {
		if !strings.Contains(slide, want) {
			t.Errorf("expected %s present, got %s", want, slide)
		}
	}
}
