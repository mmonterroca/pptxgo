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

func TestLvl1PPr_MarshalsAttrsBuFontBuCharDefRPrInOrder(t *testing.T) {
	marL, indent := bodyBulletMarL, bodyBulletIndent
	xmlStr := marshal(t, &Lvl1PPr{
		MarL:   &marL,
		Indent: &indent,
		BuFont: &drawingml.BuFont{Typeface: "Arial"},
		BuChar: &drawingml.BuChar{Char: "•"},
		DefRPr: &DefRPr{Sz: 3200},
	})

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
