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

package drawingml

import (
	"strings"
	"testing"
)

func TestGradFill_GsLstBeforeLin(t *testing.T) {
	g := &GradFill{
		RotWithShape: true,
		GsLst: &GsLst{Gs: []*Gs{
			{Pos: 0, SrgbClr: &SrgbClr{Val: "FF0000"}},
			{Pos: 100000, SrgbClr: &SrgbClr{Val: "0000FF"}},
		}},
		Lin: &Lin{Ang: 5400000, Scaled: true},
	}
	got := marshal(t, g)

	if !strings.Contains(got, `<a:gradFill rotWithShape="1">`) {
		t.Errorf("expected a:gradFill with rotWithShape=\"1\", got %s", got)
	}
	gsLstIdx := strings.Index(got, "<a:gsLst>")
	linIdx := strings.Index(got, "<a:lin ")
	if gsLstIdx == -1 || linIdx == -1 || gsLstIdx > linIdx {
		t.Errorf("expected a:gsLst before a:lin (CT_GradientFillProperties sequence), got %s", got)
	}
	if !strings.Contains(got, `<a:gs pos="0"><a:srgbClr val="FF0000">`) {
		t.Errorf("expected first gradient stop, got %s", got)
	}
	if !strings.Contains(got, `<a:gs pos="100000"><a:srgbClr val="0000FF">`) {
		t.Errorf("expected second gradient stop, got %s", got)
	}
	if !strings.Contains(got, `<a:lin ang="5400000" scaled="1">`) {
		t.Errorf("expected a:lin with ang and scaled attrs, got %s", got)
	}
}

func TestGs_SchemeClrVariant(t *testing.T) {
	g := &Gs{Pos: 50000, SchemeClr: &SchemeClr{Val: "accent1"}}
	got := marshal(t, g)

	if !strings.Contains(got, `<a:gs pos="50000"><a:schemeClr val="accent1">`) {
		t.Errorf("expected scheme-color gradient stop, got %s", got)
	}
}

func TestSrgbClr_TintShadeAlphaLumModOff(t *testing.T) {
	c := &SrgbClr{
		Val:    "FF0000",
		Tint:   &Tint{Val: 40000},
		Shade:  &Shade{Val: 20000},
		Alpha:  &Alpha{Val: 63000},
		LumMod: &LumMod{Val: 110000},
		LumOff: &LumOff{Val: 5000},
	}
	got := marshal(t, c)

	for _, want := range []string{
		`<a:srgbClr val="FF0000">`,
		`<a:tint val="40000">`,
		`<a:shade val="20000">`,
		`<a:alpha val="63000">`,
		`<a:lumMod val="110000">`,
		`<a:lumOff val="5000">`,
	} {
		if !strings.Contains(got, want) {
			t.Errorf("expected %q in %s", want, got)
		}
	}
}

func TestSchemeClr_AlphaOnlyOmitsOtherTransforms(t *testing.T) {
	// The shape a theme's own outerShdw uses (see themeFmtScheme): just alpha.
	c := &SchemeClr{Val: "accent1", Alpha: &Alpha{Val: 63000}}
	got := marshal(t, c)

	if !strings.Contains(got, `<a:schemeClr val="accent1"><a:alpha val="63000">`) {
		t.Errorf("expected schemeClr with only alpha, got %s", got)
	}
	for _, unwanted := range []string{"tint", "shade", "lumMod", "lumOff"} {
		if strings.Contains(got, unwanted) {
			t.Errorf("expected no %s element when unset, got %s", unwanted, got)
		}
	}
}
