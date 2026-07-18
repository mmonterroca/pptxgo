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
	"fmt"
	"strings"
	"testing"

	"github.com/mmonterroca/pptxgo/drawingml"
)

func TestNvPr_WithPh_OmitsIdxWhenZero(t *testing.T) {
	xmlStr := marshal(t, &NvPr{Ph: &Ph{Type: PlaceholderTitle}})
	if !strings.Contains(xmlStr, `<p:ph type="title">`) && !strings.Contains(xmlStr, `<p:ph type="title"/>`) {
		t.Errorf("expected p:ph type=\"title\" with no idx attr, got %s", xmlStr)
	}
	if strings.Contains(xmlStr, "idx=") {
		t.Errorf("idx=0 should be omitted (schema default), got %s", xmlStr)
	}
}

func TestNvPr_WithPh_EmitsExplicitNonZeroIdx(t *testing.T) {
	xmlStr := marshal(t, &NvPr{Ph: &Ph{Type: PlaceholderBody, Idx: 1}})
	if !strings.Contains(xmlStr, `idx="1"`) {
		t.Errorf("expected idx=\"1\", got %s", xmlStr)
	}
}

func TestNvPr_WithoutPh_MarshalsEmpty(t *testing.T) {
	// Regression: NvPr's new optional Ph field must not leak a p:ph into
	// ordinary (non-placeholder) shapes, pictures, or graphic frames, which
	// all construct a bare &NvPr{}.
	xmlStr := marshal(t, &NvPr{})
	if strings.Contains(xmlStr, "p:ph") {
		t.Errorf("bare NvPr must not emit p:ph, got %s", xmlStr)
	}
}

func TestLvlPPr_MarshalsAttrsBuFontBuCharDefRPrInOrder(t *testing.T) {
	marL, indent := bodyBulletMarL, bodyBulletIndent
	xmlStr := marshal(t, &LvlPPr{
		Level:  1,
		MarL:   &marL,
		Indent: &indent,
		BuFont: &drawingml.BuFont{Typeface: "Arial"},
		BuChar: &drawingml.BuChar{Char: "•"},
		DefRPr: &DefRPr{Sz: 3200},
	})

	if !strings.Contains(xmlStr, "<a:lvl1pPr ") {
		t.Errorf("expected a:lvl1pPr element name from Level=1, got %s", xmlStr)
	}
	if !strings.Contains(xmlStr, `marL="342900"`) || !strings.Contains(xmlStr, `indent="-342900"`) {
		t.Errorf("expected marL/indent attrs, got %s", xmlStr)
	}
	buFontIdx := strings.Index(xmlStr, "<a:buFont")
	buCharIdx := strings.Index(xmlStr, "<a:buChar")
	defRPrIdx := strings.Index(xmlStr, "<a:defRPr")
	if !(buFontIdx >= 0 && buCharIdx >= 0 && defRPrIdx >= 0 && buFontIdx < buCharIdx && buCharIdx < defRPrIdx) {
		t.Errorf("expected buFont < buChar < defRPr, got %s", xmlStr)
	}
}

func TestLvlPPr_MarshalsElementNamePerLevel(t *testing.T) {
	for _, lvl := range []int{1, 2, 5, 9} {
		xmlStr := marshal(t, &LvlPPr{Level: lvl, DefRPr: &DefRPr{Sz: 1800}})
		want := fmt.Sprintf("<a:lvl%dpPr>", lvl)
		if !strings.Contains(xmlStr, want) {
			t.Errorf("Level=%d: expected %s, got %s", lvl, want, xmlStr)
		}
	}
}

func TestAddShape_NvPrHasNoPlaceholderMarker(t *testing.T) {
	// Regression: an ordinary AddShape/AddTextBox shape must not carry a
	// p:ph — only Slide.AddPlaceholder (a later phase) sets one.
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(1))
	files := generateFrom(t, p)
	slide1 := string(files["ppt/slides/slide1.xml"])
	if strings.Contains(slide1, "p:ph") {
		t.Errorf("ordinary shape must not emit p:ph, got %s", slide1)
	}
}

func TestNew_MasterHasTitleAndBodyPlaceholders(t *testing.T) {
	files := generate(t)
	master := string(files["ppt/slideMasters/slideMaster1.xml"])

	if !strings.Contains(master, `<p:ph type="title">`) {
		t.Errorf("master missing title placeholder, got %s", master)
	}
	if !strings.Contains(master, `<p:ph type="body" idx="1">`) {
		t.Errorf("master missing body placeholder (type=body idx=1), got %s", master)
	}

	// Each placeholder shape must still follow CT_Shape's own
	// nvSpPr -> spPr -> txBody sequence.
	titleIdx := strings.Index(master, `name="Title Placeholder 2"`)
	if titleIdx < 0 {
		t.Fatalf("title placeholder shape not found, got %s", master)
	}
	titleSpPrIdx := strings.Index(master[titleIdx:], "<p:spPr>")
	titleTxBodyIdx := strings.Index(master[titleIdx:], "<p:txBody>")
	if !(titleSpPrIdx >= 0 && titleTxBodyIdx >= 0 && titleSpPrIdx < titleTxBodyIdx) {
		t.Errorf("title placeholder shape did not marshal nvSpPr < spPr < txBody, got %s", master)
	}
}

func TestNew_MasterBodyStyleHasAllNineLevels(t *testing.T) {
	files := generate(t)
	master := string(files["ppt/slideMasters/slideMaster1.xml"])

	bodyStyleIdx := strings.Index(master, "<p:bodyStyle>")
	bodyStyleEnd := strings.Index(master, "</p:bodyStyle>")
	if bodyStyleIdx < 0 || bodyStyleEnd < 0 {
		t.Fatalf("master missing p:bodyStyle, got %s", master)
	}
	bodyStyle := master[bodyStyleIdx:bodyStyleEnd]

	prevIdx := -1
	for lvl := 1; lvl <= 9; lvl++ {
		want := fmt.Sprintf("<a:lvl%dpPr", lvl)
		idx := strings.Index(bodyStyle, want)
		if idx < 0 {
			t.Fatalf("missing %s in bodyStyle, got %s", want, bodyStyle)
		}
		if idx < prevIdx {
			t.Errorf("expected levels in ascending order, %s appeared before the previous level", want)
		}
		prevIdx = idx
	}
	// titleStyle/otherStyle only need their first level.
	titleStyleIdx := strings.Index(master, "<p:titleStyle>")
	titleStyleEnd := strings.Index(master, "</p:titleStyle>")
	titleStyle := master[titleStyleIdx:titleStyleEnd]
	if strings.Contains(titleStyle, "lvl2pPr") {
		t.Errorf("expected titleStyle to keep only its first level, got %s", titleStyle)
	}
}

func TestNew_MasterBodyStyleLevelsIncreaseIndentAndAlternateGlyph(t *testing.T) {
	files := generate(t)
	master := string(files["ppt/slideMasters/slideMaster1.xml"])

	if !strings.Contains(master, `<a:lvl1pPr marL="342900" indent="-342900">`) {
		t.Errorf("expected level 1 marL=342900, got %s", master)
	}
	if !strings.Contains(master, `<a:lvl2pPr marL="685800" indent="-342900">`) {
		t.Errorf("expected level 2 marL=685800 (double level 1's), got %s", master)
	}
	lvl1Idx := strings.Index(master, "<a:lvl1pPr")
	lvl2Idx := strings.Index(master, "<a:lvl2pPr")
	bullet1 := master[lvl1Idx:lvl2Idx]
	if !strings.Contains(bullet1, `<a:buChar char="•">`) {
		t.Errorf("expected level 1 bullet glyph •, got %s", bullet1)
	}
	lvl3Idx := strings.Index(master, "<a:lvl3pPr")
	bullet2 := master[lvl2Idx:lvl3Idx]
	if !strings.Contains(bullet2, `<a:buChar char="–">`) {
		t.Errorf("expected level 2 bullet glyph – (alternating from level 1), got %s", bullet2)
	}
}

func TestNew_MasterBodyStyleHasBulletDefaults(t *testing.T) {
	files := generate(t)
	master := string(files["ppt/slideMasters/slideMaster1.xml"])

	bodyStyleIdx := strings.Index(master, "<p:bodyStyle>")
	if bodyStyleIdx < 0 {
		t.Fatalf("master missing p:bodyStyle, got %s", master)
	}
	buFontIdx := strings.Index(master[bodyStyleIdx:], "<a:buFont")
	buCharIdx := strings.Index(master[bodyStyleIdx:], "<a:buChar")
	defRPrIdx := strings.Index(master[bodyStyleIdx:], "<a:defRPr")
	if !(buFontIdx >= 0 && buCharIdx >= 0 && defRPrIdx >= 0 && buFontIdx < buCharIdx && buCharIdx < defRPrIdx) {
		t.Errorf("bodyStyle's lvl1pPr must marshal buFont < buChar < defRPr, got %s", master[bodyStyleIdx:bodyStyleIdx+300])
	}
}

func TestNew_MasterPlaceholdersHaveExplicitRectGeometry(t *testing.T) {
	// The master's own placeholders are the inheritance chain's root — they
	// have no ancestor to inherit a:prstGeom from — so unlike a layout/slide
	// placeholder (which legitimately omits it), they must declare an
	// explicit rect geometry for viewers that don't default a geometry-less
	// placeholder to a rectangle.
	files := generate(t)
	master := string(files["ppt/slideMasters/slideMaster1.xml"])

	titleIdx := strings.Index(master, `name="Title Placeholder 2"`)
	bodyIdx := strings.Index(master, `name="Body Placeholder 3"`)
	if titleIdx < 0 || bodyIdx < 0 {
		t.Fatalf("master placeholder shapes not found, got %s", master)
	}
	if !strings.Contains(master[titleIdx:bodyIdx], `<a:prstGeom prst="rect">`) {
		t.Errorf("master title placeholder missing explicit rect prstGeom, got %s", master[titleIdx:bodyIdx])
	}
	if !strings.Contains(master[bodyIdx:], `<a:prstGeom prst="rect">`) {
		t.Errorf("master body placeholder missing explicit rect prstGeom, got %s", master[bodyIdx:])
	}
}

func TestNew_MasterPlaceholderGeometryScalesWithSlideSize(t *testing.T) {
	// The 16:9 default and 4:3 preset share the same slide height
	// (6858000 EMU) but differ in width, so the title placeholder's width
	// must differ between them rather than a hardcoded 16:9 value leaking
	// into a differently-sized presentation.
	widescreen := generate(t)
	fourByThree := generateFrom(t, New(WithStandard4x3()))

	wideMaster := string(widescreen["ppt/slideMasters/slideMaster1.xml"])
	fourByThreeMaster := string(fourByThree["ppt/slideMasters/slideMaster1.xml"])

	if !strings.Contains(wideMaster, `cx="10972800"`) {
		t.Errorf("16:9 title width should be 10972800 (12192000 - 2*5%%), got %s", wideMaster)
	}
	if !strings.Contains(fourByThreeMaster, `cx="8229600"`) {
		t.Errorf("4:3 title width should be 8229600 (9144000 - 2*5%%), got %s", fourByThreeMaster)
	}
}
