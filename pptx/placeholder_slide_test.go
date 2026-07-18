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

func TestAddSlide_NoOptionsDefaultsToBlankLayout(t *testing.T) {
	// Regression: every pre-existing p.AddSlide() call (no args) must keep
	// targeting slideLayout1.xml, the blank layout — WithLayout's default
	// must not silently change what an existing caller renders.
	files := generate(t)
	rels := string(files["ppt/slides/_rels/slide1.xml.rels"])
	if !strings.Contains(rels, "slideLayout1.xml") {
		t.Errorf("expected slide1 to target slideLayout1.xml by default, got %s", rels)
	}
}

func TestAddSlide_WithLayoutSelectsMatchingLayoutPart(t *testing.T) {
	p := New()
	p.AddSlide()                                  // slide1 -> slideLayout1 (blank)
	p.AddSlide(WithLayout(LayoutTitleAndContent)) // slide2 -> slideLayout3 (see newStandardLayouts order)
	p.AddSlide(WithLayout(LayoutTwoContent))      // slide3 -> slideLayout5
	files := generateFrom(t, p)

	rels2 := string(files["ppt/slides/_rels/slide2.xml.rels"])
	if !strings.Contains(rels2, SlideLayoutPath(3)[len("ppt/slideLayouts/"):]) {
		t.Errorf("expected slide2 to target %s, got %s", SlideLayoutPath(3), rels2)
	}

	rels3 := string(files["ppt/slides/_rels/slide3.xml.rels"])
	if !strings.Contains(rels3, SlideLayoutPath(5)[len("ppt/slideLayouts/"):]) {
		t.Errorf("expected slide3 to target %s, got %s", SlideLayoutPath(5), rels3)
	}
}

func TestAddSlide_UnknownLayoutAccumulatesError(t *testing.T) {
	p := New()
	p.AddSlide(WithLayout(LayoutType("noSuchLayout")))

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated unknown-layout error")
	}
}

func TestSlide_AddPlaceholder_OmitsXfrmToInherit(t *testing.T) {
	p := New()
	s := p.AddSlide(WithLayout(LayoutTitleAndContent))
	s.AddPlaceholder(PlaceholderTitle, 0)
	files := generateFrom(t, p)

	slide1 := string(files["ppt/slides/slide1.xml"])
	if strings.Contains(slide1, "<a:xfrm>") {
		t.Errorf("a slide placeholder should omit a:xfrm to inherit from its layout, got %s", slide1)
	}
	if !strings.Contains(slide1, `<p:ph type="title">`) && !strings.Contains(slide1, `<p:ph type="title"/>`) {
		t.Errorf("expected a title placeholder, got %s", slide1)
	}
}

func TestSlide_Title_EmitsTitlePlaceholderWithText(t *testing.T) {
	p := New()
	s := p.AddSlide(WithLayout(LayoutTitleAndContent))
	s.Title("Next Steps").Bold()
	files := generateFrom(t, p)

	slide1 := string(files["ppt/slides/slide1.xml"])
	if !strings.Contains(slide1, `<p:ph type="title">`) && !strings.Contains(slide1, `<p:ph type="title"/>`) {
		t.Errorf("expected a title placeholder, got %s", slide1)
	}
	if !strings.Contains(slide1, "Next Steps") {
		t.Errorf("expected title text, got %s", slide1)
	}
	// Title returns a *Paragraph, so further formatting (Bold here) must
	// still apply to the run Text started.
	if !strings.Contains(slide1, `b="1"`) {
		t.Errorf("expected Bold to mark the title run, got %s", slide1)
	}
}

func TestSlide_Body_EmitsBodyPlaceholderIdxOneWithText(t *testing.T) {
	p := New()
	s := p.AddSlide(WithLayout(LayoutTitleAndContent))
	s.Body("Renew the partner agreement")
	files := generateFrom(t, p)

	slide1 := string(files["ppt/slides/slide1.xml"])
	if !strings.Contains(slide1, `<p:ph type="body" idx="1">`) {
		t.Errorf("expected a body placeholder (idx=1), got %s", slide1)
	}
	if !strings.Contains(slide1, "Renew the partner agreement") {
		t.Errorf("expected body text, got %s", slide1)
	}
}

func TestSlide_AddPlaceholder_ShapeIDsStaySequentialWithOtherShapes(t *testing.T) {
	// Regression: AddPlaceholder must share the slide's own nextShapeID
	// counter with AddShape/AddTextBox/AddImage/AddTable, not its own —
	// otherwise mixing placeholder and freeform shapes on one slide could
	// produce duplicate p:cNvPr ids.
	p := New()
	s := p.AddSlide(WithLayout(LayoutTitleAndContent))
	s.AddPlaceholder(PlaceholderTitle, 0)                             // id 2
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(1), Inches(1)) // id 3
	s.AddPlaceholder(PlaceholderBody, 1)                              // id 4
	files := generateFrom(t, p)

	slide1 := string(files["ppt/slides/slide1.xml"])
	for _, id := range []string{`id="2"`, `id="3"`, `id="4"`} {
		if !strings.Contains(slide1, id) {
			t.Errorf("expected shape %s, got %s", id, slide1)
		}
	}
}
