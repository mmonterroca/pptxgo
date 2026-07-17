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
	"encoding/xml"
	"strings"
	"testing"
)

func marshal(t *testing.T, v any) string {
	t.Helper()
	b, err := xml.Marshal(v)
	if err != nil {
		t.Fatalf("xml.Marshal: %v", err)
	}
	return string(b)
}

func TestXfrm_MarshalsWithAPrefix(t *testing.T) {
	x := &Xfrm{Off: &Off{X: 100, Y: 200}, Ext: &Ext{Cx: 300, Cy: 400}}
	got := marshal(t, x)

	for _, want := range []string{
		`<a:xfrm>`,
		`<a:off x="100" y="200">`,
		`<a:ext cx="300" cy="400">`,
	} {
		if !strings.Contains(got, want) {
			t.Errorf("expected %q in %s", want, got)
		}
	}
	// Zero rotation and no flips must not appear (omitempty).
	if strings.Contains(got, `rot=`) || strings.Contains(got, `flipH`) || strings.Contains(got, `flipV`) {
		t.Errorf("expected zero-value attrs omitted, got %s", got)
	}
}

func TestXfrm_RotationAndFlipsRoundTrip(t *testing.T) {
	x := &Xfrm{Rot: 5400000, FlipH: true, Ext: &Ext{Cx: 1, Cy: 1}}
	got := marshal(t, x)
	if !strings.Contains(got, `rot="5400000"`) {
		t.Errorf("expected rot attr, got %s", got)
	}
	if !strings.Contains(got, `flipH="true"`) {
		t.Errorf("expected flipH attr, got %s", got)
	}
}

func TestPrstGeom_DefaultHasNoAvLst(t *testing.T) {
	g := &PrstGeom{Prst: "rect"}
	got := marshal(t, g)
	if !strings.Contains(got, `prst="rect"`) {
		t.Errorf("expected prst attr, got %s", got)
	}
	if strings.Contains(got, "avLst") {
		t.Errorf("expected no avLst element when unset, got %s", got)
	}
}

func TestBlip_CarriesOwnNamespace(t *testing.T) {
	b := NewBlip("rId3")
	got := marshal(t, b)
	if !strings.Contains(got, `xmlns:r="`+NamespaceRelationships+`"`) {
		t.Errorf("expected self-contained xmlns:r declaration, got %s", got)
	}
	if !strings.Contains(got, `r:embed="rId3"`) {
		t.Errorf("expected r:embed attr, got %s", got)
	}
}

func TestSolidFill_RGBVsScheme(t *testing.T) {
	rgb := NewSolidFillRGB(Color{R: 0xFF, G: 0, B: 0})
	got := marshal(t, rgb)
	if !strings.Contains(got, `<a:srgbClr val="FF0000">`) {
		t.Errorf("expected srgbClr, got %s", got)
	}
	if strings.Contains(got, "schemeClr") {
		t.Errorf("did not expect schemeClr in RGB fill, got %s", got)
	}

	scheme := NewSolidFillScheme("accent1")
	got = marshal(t, scheme)
	if !strings.Contains(got, `<a:schemeClr val="accent1">`) {
		t.Errorf("expected schemeClr, got %s", got)
	}
}

func TestColor_HexRoundTrip(t *testing.T) {
	cases := []struct {
		hex  string
		want Color
	}{
		{"FF0000", Color{R: 255, G: 0, B: 0}},
		{"#00FF00", Color{R: 0, G: 255, B: 0}},
		{"00F", Color{R: 0, G: 0, B: 255}},
	}
	for _, tc := range cases {
		got, err := FromHex(tc.hex)
		if err != nil {
			t.Fatalf("FromHex(%q): %v", tc.hex, err)
		}
		if got != tc.want {
			t.Errorf("FromHex(%q) = %+v, want %+v", tc.hex, got, tc.want)
		}
	}

	if got := ToHex(Color{R: 18, G: 52, B: 86}); got != "123456" {
		t.Errorf("ToHex = %s, want 123456", got)
	}
}

func TestColor_FromHexRejectsInvalidLength(t *testing.T) {
	if _, err := FromHex("12345"); err == nil {
		t.Error("expected error for 5-character hex string")
	}
}

func TestGraphicData_WrapsArbitraryInnerContent(t *testing.T) {
	type stubTable struct {
		XMLName xml.Name `xml:"a:tbl"`
	}
	g := NewGraphic(&GraphicData{URI: GraphicDataURITable, Inner: &stubTable{}})
	got := marshal(t, g)
	if !strings.Contains(got, `uri="`+GraphicDataURITable+`"`) {
		t.Errorf("expected table URI, got %s", got)
	}
	if !strings.Contains(got, `<a:tbl>`) {
		t.Errorf("expected inner content marshaled, got %s", got)
	}
}

func TestEMUsPerPoint_ConsistentWithEMUsPerInch(t *testing.T) {
	if PointsPerInch*EMUsPerPoint != EMUsPerInch {
		t.Fatalf("EMUsPerPoint (%d) inconsistent with EMUsPerInch (%d) at %d points/inch",
			EMUsPerPoint, EMUsPerInch, PointsPerInch)
	}
}
