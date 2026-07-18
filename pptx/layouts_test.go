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
	"strconv"
	"strings"
	"testing"
)

func TestNew_ProducesAllStandardLayoutParts(t *testing.T) {
	files := generate(t)
	for n := 1; n <= 5; n++ {
		path := SlideLayoutPath(n)
		if _, ok := files[path]; !ok {
			t.Errorf("missing %s", path)
		}
		rels := "ppt/slideLayouts/_rels/slideLayout" + strconv.Itoa(n) + ".xml.rels"
		if _, ok := files[rels]; !ok {
			t.Errorf("missing %s", rels)
		}
	}
}

func TestNew_MasterHasFiveLayoutEntries(t *testing.T) {
	files := generate(t)
	master := string(files["ppt/slideMasters/slideMaster1.xml"])
	count := strings.Count(master, "<p:sldLayoutId ")
	if count != 5 {
		t.Errorf("expected 5 p:sldLayoutId entries, got %d in %s", count, master)
	}
}

func TestNew_Slide1StillUsesBlankLayout(t *testing.T) {
	// Regression: AddSlide's hardcoded rel target is slideLayout1.xml, so
	// that part must stay the Blank layout (no placeholders) — otherwise
	// every existing caller's slide would suddenly render title/body
	// placeholders it never asked for.
	files := generate(t)
	slideRels := string(files["ppt/slides/_rels/slide1.xml.rels"])
	if !strings.Contains(slideRels, "slideLayout1.xml") {
		t.Fatalf("slide1 should still reference slideLayout1.xml, got %s", slideRels)
	}
	layout1 := string(files["ppt/slideLayouts/slideLayout1.xml"])
	if strings.Contains(layout1, "p:ph") {
		t.Errorf("slideLayout1.xml must stay the blank (placeholder-free) layout, got %s", layout1)
	}
	if !strings.Contains(layout1, `type="blank"`) {
		t.Errorf("slideLayout1.xml should carry type=\"blank\", got %s", layout1)
	}
}

func TestNew_TitleSlideLayoutPlaceholdersHaveOwnGeometry(t *testing.T) {
	// ctrTitle/subTitle share no type+idx with any master placeholder, so
	// neither can inherit position — both must carry their own a:xfrm.
	files := generate(t)
	layout := string(files[SlideLayoutPath(2)])
	if !strings.Contains(layout, `type="title"`) {
		t.Fatalf("slideLayout2 should be the Title Slide layout, got %s", layout)
	}
	if !strings.Contains(layout, `<p:ph type="ctrTitle">`) && !strings.Contains(layout, `<p:ph type="ctrTitle"/>`) {
		t.Errorf("missing ctrTitle placeholder, got %s", layout)
	}
	if !strings.Contains(layout, `<p:ph type="subTitle" idx="1">`) {
		t.Errorf("missing subTitle placeholder (idx=1), got %s", layout)
	}
	if strings.Count(layout, "<a:xfrm>") != 2 {
		t.Errorf("both ctrTitle and subTitle should declare their own a:xfrm, got %s", layout)
	}
}

func TestNew_TitleAndContentLayoutPlaceholdersInheritFromMaster(t *testing.T) {
	// title (idx=0) and body (idx=1) match the master's own placeholders,
	// so this layout can legitimately omit a:xfrm and inherit position.
	files := generate(t)
	layout := string(files[SlideLayoutPath(3)])
	if !strings.Contains(layout, `type="obj"`) {
		t.Fatalf("slideLayout3 should be the Title and Content layout, got %s", layout)
	}
	if strings.Contains(layout, "<a:xfrm>") {
		t.Errorf("Title and Content placeholders should inherit geometry (no a:xfrm), got %s", layout)
	}
	if !strings.Contains(layout, `<p:ph type="title">`) && !strings.Contains(layout, `<p:ph type="title"/>`) {
		t.Errorf("missing title placeholder, got %s", layout)
	}
	if !strings.Contains(layout, `<p:ph type="body" idx="1">`) {
		t.Errorf("missing body placeholder (idx=1), got %s", layout)
	}
}

func TestNew_TwoContentLayoutBodiesAreSideBySideNotOverlapping(t *testing.T) {
	files := generate(t)
	layout := string(files[SlideLayoutPath(5)])
	if !strings.Contains(layout, `type="twoObj"`) {
		t.Fatalf("slideLayout5 should be the Two Content layout, got %s", layout)
	}
	if !strings.Contains(layout, `<p:ph type="body" idx="1">`) || !strings.Contains(layout, `<p:ph type="body" idx="2">`) {
		t.Errorf("expected two distinctly-indexed body placeholders, got %s", layout)
	}

	leftOff, leftExt := xfrmAfter(t, layout, `idx="1"`)
	rightOff, _ := xfrmAfter(t, layout, `idx="2"`)
	if leftOff+leftExt > rightOff {
		t.Errorf("left body (x=%d, w=%d) overlaps right body (x=%d)", leftOff, leftExt, rightOff)
	}
}

// xfrmAfter finds the first a:off/x and a:ext/cx attribute values in xmlStr
// after the given marker substring — a small scoped-parse helper for
// assertions that need actual coordinate values rather than mere presence.
func xfrmAfter(t *testing.T, xmlStr, marker string) (offX, extCx int) {
	t.Helper()
	i := strings.Index(xmlStr, marker)
	if i < 0 {
		t.Fatalf("marker %q not found in %s", marker, xmlStr)
	}
	rest := xmlStr[i:]
	offX = intAttrAfter(t, rest, `<a:off x="`)
	extCx = intAttrAfter(t, rest, `<a:ext cx="`)
	return offX, extCx
}

func intAttrAfter(t *testing.T, xmlStr, marker string) int {
	t.Helper()
	i := strings.Index(xmlStr, marker)
	if i < 0 {
		t.Fatalf("marker %q not found in %s", marker, xmlStr)
	}
	rest := xmlStr[i+len(marker):]
	end := strings.Index(rest, `"`)
	if end < 0 {
		t.Fatalf("unterminated attribute after %q", marker)
	}
	n, err := strconv.Atoi(rest[:end])
	if err != nil {
		t.Fatalf("parse int after %q: %v", marker, err)
	}
	return n
}

func TestNew_LayoutGeometryScalesWithSlideSize(t *testing.T) {
	wide := generate(t)
	fourByThree := generateFrom(t, New(WithStandard4x3()))

	wideLayout := string(wide[SlideLayoutPath(2)])
	fourByThreeLayout := string(fourByThree[SlideLayoutPath(2)])

	if !strings.Contains(wideLayout, `cx="10972800"`) {
		t.Errorf("16:9 ctrTitle width should be 10972800, got %s", wideLayout)
	}
	if !strings.Contains(fourByThreeLayout, `cx="8229600"`) {
		t.Errorf("4:3 ctrTitle width should be 8229600, got %s", fourByThreeLayout)
	}
}
