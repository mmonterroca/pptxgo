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

// --- table cell per-side borders --------------------------------------------

func TestTableCell_BorderEmitsNamedSideWithinTcPr(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tbl := s.AddTable(1, 1, Inches(1), Inches(1), Inches(2), Inches(1))
	tbl.Cell(0, 0).Border(SideTop, RGB(0x1F, 0x49, 0x7D), 1.5)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:lnT w="19050">`) {
		t.Errorf("expected a:lnT with width, got %s", slide)
	}
	if !strings.Contains(slide, `<a:srgbClr val="1F497D">`) {
		t.Errorf("expected the border's color, got %s", slide)
	}
	for _, unwanted := range []string{"a:lnL", "a:lnR", "a:lnB", "a:lnTlToBr", "a:lnBlToTr"} {
		if strings.Contains(slide, unwanted) {
			t.Errorf("expected only the requested side, got %s", slide)
		}
	}
}

func TestTableCell_BorderSchemeEmitsSchemeClr(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tbl := s.AddTable(1, 1, Inches(1), Inches(1), Inches(2), Inches(1))
	tbl.Cell(0, 0).BorderScheme(SideDiagonalDown, SchemeAccent1, 1.0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:lnTlToBr w="12700">`) {
		t.Errorf("expected a:lnTlToBr with width, got %s", slide)
	}
	if !strings.Contains(slide, `<a:schemeClr val="accent1">`) {
		t.Errorf("expected the border's schemeClr, got %s", slide)
	}
}

func TestTableCell_BorderMultipleSidesCoexist(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tbl := s.AddTable(1, 1, Inches(1), Inches(1), Inches(2), Inches(1))
	tbl.Cell(0, 0).
		Border(SideTop, RGB(0, 0, 0), 1.0).
		Border(SideBottom, RGB(0xFF, 0, 0), 1.0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, "<a:lnT ") || !strings.Contains(slide, "<a:lnB ") {
		t.Errorf("expected both a:lnT and a:lnB present, got %s", slide)
	}
}

func TestTableCell_BorderInvalidSideAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tbl := s.AddTable(1, 1, Inches(1), Inches(1), Inches(2), Inches(1))
	tbl.Cell(0, 0).Border(TableCellSide("bogus"), RGB(0, 0, 0), 1.0)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated invalid-side error")
	}
}

func TestTableCell_BorderOutOfRangeWidthAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tbl := s.AddTable(1, 1, Inches(1), Inches(1), Inches(2), Inches(1))
	tbl.Cell(0, 0).Border(SideTop, RGB(0, 0, 0), -1)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated border-width error")
	}
}

// --- gradient stop tint/shade ------------------------------------------------

func TestGradientFill_StopTintShadeEmitTintShadeElements(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		GradientFill(90,
			GradientStop{Scheme: SchemeAccent2, Pos: 0, Shade: 20},
			GradientStop{Scheme: SchemeAccent4, Pos: 100, Tint: 30})

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:schemeClr val="accent2">`) || !strings.Contains(slide, `<a:shade val="20000">`) {
		t.Errorf("expected first stop's schemeClr and shade, got %s", slide)
	}
	if !strings.Contains(slide, `<a:schemeClr val="accent4">`) || !strings.Contains(slide, `<a:tint val="30000">`) {
		t.Errorf("expected second stop's schemeClr and tint, got %s", slide)
	}
}

func TestGradientFill_StopTintShadeOutOfRangeAccumulatesError(t *testing.T) {
	for _, bad := range []GradientStop{
		{Color: RGB(0, 0, 0), Pos: 0, Tint: 101},
		{Color: RGB(0, 0, 0), Pos: 0, Shade: -1},
	} {
		p := New()
		s := p.AddSlide()
		s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
			GradientFill(0, bad, GradientStop{Color: RGB(0xFF, 0xFF, 0xFF), Pos: 100})
		if err := p.Save(&bytes.Buffer{}); err == nil {
			t.Errorf("expected Save to error for out-of-range stop %+v", bad)
		}
	}
}

func TestGradientFill_StopWithBothTintAndShadeAccumulatesError(t *testing.T) {
	// Regression: Tint and Shade together lighten-then-darken to a muddied
	// color; the documented "at most one" contract must be enforced, not
	// silently emitted as both a:tint and a:shade.
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		GradientFill(0,
			GradientStop{Scheme: SchemeAccent2, Pos: 0, Tint: 30, Shade: 20},
			GradientStop{Color: RGB(0xFF, 0xFF, 0xFF), Pos: 100})

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated both-Tint-and-Shade error")
	}
}

// --- line cap / join / arrowheads --------------------------------------------

func TestLineCap_SetsAttrAfterBorder(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		Border(RGB(0, 0, 0), 2).
		LineCap(LineCapRound)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `cap="rnd"`) {
		t.Errorf("expected cap=\"rnd\", got %s", slide)
	}
}

func TestLineCap_BeforeBorderAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).LineCap(LineCapRound)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated no-outline error")
	}
}

func TestLineJoin_MiterUsesOfficeDefaultLimit(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		Border(RGB(0, 0, 0), 2).
		LineJoin(LineJoinMiter)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:miter lim="800000">`) {
		t.Errorf("expected a:miter lim=\"800000\", got %s", slide)
	}
}

func TestLineJoin_SwitchingStylesClearsThePrevious(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		Border(RGB(0, 0, 0), 2).
		LineJoin(LineJoinMiter).
		LineJoin(LineJoinRound)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if strings.Contains(slide, "a:miter") {
		t.Errorf("expected the second LineJoin to clear the first, got %s", slide)
	}
	if !strings.Contains(slide, "<a:round></a:round>") {
		t.Errorf("expected a:round present, got %s", slide)
	}
}

func TestArrowStartEnd_EmitDistinctHeadAndTailEndTypes(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeLine, Inches(1), Inches(1), Emu(1), Inches(2)).
		Border(RGB(0, 0, 0), 1.5).
		ArrowStart(ArrowheadOval).
		ArrowEnd(ArrowheadTriangle)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:headEnd type="oval" w="med" len="med">`) {
		t.Errorf("expected headEnd type=oval, got %s", slide)
	}
	if !strings.Contains(slide, `<a:tailEnd type="triangle" w="med" len="med">`) {
		t.Errorf("expected tailEnd type=triangle, got %s", slide)
	}
}

func TestArrowEnd_BeforeBorderAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeLine, Inches(1), Inches(1), Emu(1), Inches(2)).ArrowEnd(ArrowheadTriangle)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated no-outline error")
	}
}

func TestArrowEnd_InvalidTypeAccumulatesError(t *testing.T) {
	// Regression: an unrecognized ST_LineEndType must be rejected (like
	// LineCap's own enum check), not written straight into a schema-invalid
	// a:tailEnd/@type.
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeLine, Inches(1), Inches(1), Emu(1), Inches(2)).
		Border(RGB(0, 0, 0), 1).
		ArrowEnd(ArrowheadType("arrowhead"))

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated invalid-arrowhead-type error")
	}
}

func TestLineJoin_InvalidStylePreservesPriorJoin(t *testing.T) {
	// Regression: an invalid style must not wipe a previously-set join —
	// LineJoin validates before clearing, matching LineCap. The accumulated
	// error gates Save, so the surviving join is asserted on the in-memory
	// struct directly rather than through the emitted XML.
	p := New()
	s := p.AddSlide()
	sh := s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		Border(RGB(0, 0, 0), 2).
		LineJoin(LineJoinRound).
		LineJoin(LineJoinStyle("bad"))

	if sh.spPr.Ln.Round == nil {
		t.Error("expected the prior round join to survive an invalid LineJoin call")
	}
	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Error("expected the invalid LineJoin to still record an error on Save")
	}
}

// --- shape effects (shadow / glow / reflection / soft edges) ----------------

func TestShadow_EmitsOuterShdwWithAlphaColor(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		Shadow(RGB(0, 0, 0), 60)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `rotWithShape="0"`) {
		t.Errorf("expected rotWithShape=\"0\" (Office's own preset), got %s", slide)
	}
	if !strings.Contains(slide, `<a:srgbClr val="000000">`) || !strings.Contains(slide, `<a:alpha val="60000">`) {
		t.Errorf("expected shadow color with alpha, got %s", slide)
	}
}

func TestShadow_OutOfRangeAlphaAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).Shadow(RGB(0, 0, 0), 101)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated out-of-range alpha error")
	}
}

func TestGlow_EmitsRadiusInEMUs(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		Glow(RGB(0xED, 0x7D, 0x31), 5)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:glow rad="63500">`) {
		t.Errorf("expected glow rad=\"63500\" (5pt in EMUs), got %s", slide)
	}
}

func TestReflection_EmitsStartOpacity(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).Reflection(50)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `stA="50000"`) {
		t.Errorf("expected stA=\"50000\", got %s", slide)
	}
	if !strings.Contains(slide, `sy="-100000"`) {
		t.Errorf("expected the mirror-flip sy=\"-100000\" (without it nothing renders), got %s", slide)
	}
}

func TestReflection_OutOfRangeAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).Reflection(-1)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated out-of-range error")
	}
}

func TestSoftEdges_EmitsRadiusInEMUs(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).SoftEdges(2)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:softEdge rad="25400">`) {
		t.Errorf("expected softEdge rad=\"25400\" (2pt in EMUs), got %s", slide)
	}
}

func TestSoftEdges_NegativeRadiusAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).SoftEdges(-1)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated negative-radius error")
	}
}

func TestPictureRef_EffectsShareShapeRefHelpers(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddImageFromBytes(pngBytes(t, 40, 40), Inches(1), Inches(1)).
		Shadow(RGB(0, 0, 0), 50).
		Glow(RGB(0xFF, 0, 0), 3).
		Reflection(40).
		SoftEdges(1)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	for _, want := range []string{"<a:outerShdw", "<a:glow", "<a:reflection", "<a:softEdge"} {
		if !strings.Contains(slide, want) {
			t.Errorf("expected %q on the picture, got %s", want, slide)
		}
	}
}
