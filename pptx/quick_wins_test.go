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

// --- gradient scheme-color stops -------------------------------------------

func TestGradientFill_SchemeStopEmitsSchemeClrNotSrgbClr(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		GradientFill(90,
			GradientStop{Scheme: SchemeAccent1, Pos: 0},
			GradientStop{Scheme: SchemeAccent2, Pos: 100})

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, "<a:gradFill") {
		t.Fatalf("expected a:gradFill, got %s", slide)
	}
	for _, want := range []string{`<a:schemeClr val="accent1">`, `<a:schemeClr val="accent2">`} {
		if !strings.Contains(slide, want) {
			t.Errorf("expected gradient stop %q, got %s", want, slide)
		}
	}
	if strings.Contains(slide, "a:srgbClr") {
		t.Errorf("expected no a:srgbClr for scheme-colored stops, got %s", slide)
	}
}

func TestGradientFill_MixedSchemeAndRgbStops(t *testing.T) {
	// Scheme takes precedence per stop; a stop with no Scheme uses its RGB.
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		GradientFill(0,
			GradientStop{Scheme: SchemeAccent1, Pos: 0},
			GradientStop{Color: RGB(0xFF, 0x00, 0x00), Pos: 100})

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:schemeClr val="accent1">`) {
		t.Errorf("expected scheme stop accent1, got %s", slide)
	}
	if !strings.Contains(slide, `<a:srgbClr val="FF0000">`) {
		t.Errorf("expected RGB stop FF0000, got %s", slide)
	}
}

// --- numbered bullet StartAt -----------------------------------------------

func TestNumberedBulletFrom_EmitsStartAt(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("Item four").NumberedBulletFrom(NumArabicPeriod, 4)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:buAutoNum type="arabicPeriod" startAt="4">`) {
		t.Errorf("expected buAutoNum with startAt=\"4\", got %s", slide)
	}
}

func TestNumberedBulletFrom_OutOfRangeAccumulatesError(t *testing.T) {
	for _, bad := range []int{0, -1, 40000} {
		p := New()
		s := p.AddSlide()
		tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
		tb.AddParagraph().Text("x").NumberedBulletFrom(NumArabicPeriod, bad)
		if err := p.Save(&bytes.Buffer{}); err == nil {
			t.Errorf("expected Save to error for startAt=%d", bad)
		}
	}
}

// --- shape adjust handles (AvLst) ------------------------------------------

func TestAdjust_AppendsGuideWithValFormula(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRoundRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		Adjust("adj", 25000)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:gd name="adj" fmla="val 25000">`) {
		t.Errorf("expected an adjust guide gd name=\"adj\" fmla=\"val 25000\", got %s", slide)
	}
}

func TestAdjust_SameNameOverwritesRatherThanDuplicates(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRoundRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		Adjust("adj", 10000).
		Adjust("adj", 30000)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if strings.Count(slide, `name="adj"`) != 1 {
		t.Errorf("expected a single adj guide after overwrite, got %s", slide)
	}
	if !strings.Contains(slide, `fmla="val 30000"`) || strings.Contains(slide, `fmla="val 10000"`) {
		t.Errorf("expected the second Adjust to overwrite the first, got %s", slide)
	}
}

func TestAdjust_OnInheritedGeometryPlaceholderAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide(WithLayout(LayoutTitleAndContent))
	// A slide placeholder inherits its geometry and has no prstGeom of its own.
	s.AddPlaceholder(PlaceholderBody, 1).Adjust("adj", 25000)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated Adjust-on-placeholder error")
	}
}

// --- table cell properties (TcPr) ------------------------------------------

func TestTableCell_FillSchemeEmitsTcPrSolidFillSchemeClr(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tbl := s.AddTable(2, 2, Inches(1), Inches(1), Inches(6), Inches(2))
	header := tbl.Cell(0, 0)
	header.Text("Header")
	header.FillScheme(SchemeAccent1)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, "<a:tcPr") {
		t.Fatalf("expected a:tcPr, got %s", slide)
	}
	if !strings.Contains(slide, `<a:tcPr><a:solidFill><a:schemeClr val="accent1">`) &&
		!strings.Contains(slide, `<a:schemeClr val="accent1">`) {
		t.Errorf("expected cell fill schemeClr accent1 inside tcPr, got %s", slide)
	}
}

func TestTableCell_TcPrOrderTxBodyBeforeTcPr(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tbl := s.AddTable(1, 1, Inches(1), Inches(1), Inches(2), Inches(1))
	cell := tbl.Cell(0, 0)
	cell.Text("Hi")
	cell.Anchor(AnchorMiddle).Fill(RGB(0x1F, 0x49, 0x7D))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	txIdx := strings.Index(slide, "<a:txBody>")
	tcPrIdx := strings.Index(slide, "<a:tcPr")
	if txIdx == -1 || tcPrIdx == -1 || txIdx > tcPrIdx {
		t.Errorf("expected a:txBody before a:tcPr (schema order), got %s", slide)
	}
	if !strings.Contains(slide, `anchor="ctr"`) {
		t.Errorf("expected cell anchor=\"ctr\", got %s", slide)
	}
	if !strings.Contains(slide, `<a:srgbClr val="1F497D">`) {
		t.Errorf("expected cell fill 1F497D, got %s", slide)
	}
}

func TestTableCell_NoFillEmitsNoFillInsideTcPr(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tbl := s.AddTable(1, 1, Inches(1), Inches(1), Inches(2), Inches(1))
	tbl.Cell(0, 0).Fill(RGB(0xFF, 0, 0))
	tbl.Cell(0, 0).NoFill()

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, "<a:tcPr><a:noFill>") && !strings.Contains(slide, "<a:noFill>") {
		t.Errorf("expected a:noFill inside tcPr, got %s", slide)
	}
	if strings.Contains(slide, "a:solidFill") {
		t.Errorf("expected NoFill to clear the earlier Fill on the cell, got %s", slide)
	}
}
