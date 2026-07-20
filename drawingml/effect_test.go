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

func TestReflection_EmitsMirrorScaleSy(t *testing.T) {
	// Regression: a reflection without sy=-100000 is a copy that is NOT
	// flipped, so nothing visible renders in any viewer even though the XML
	// validates. sy is the essential attribute.
	r := &Reflection{BlurRad: 6350, StA: 35000, EndPos: 55000, Dir: 5400000, Sy: -100000, Algn: "bl"}
	got := marshal(t, r)
	if !strings.Contains(got, `sy="-100000"`) {
		t.Errorf("expected sy=\"-100000\" (the mirror flip), got %s", got)
	}
	if !strings.Contains(got, `algn="bl"`) {
		t.Errorf("expected algn=\"bl\", got %s", got)
	}
}

func TestEffectLst_ChildOrderMirrorsSchemaSequence(t *testing.T) {
	e := &EffectLst{
		Glow:       &Glow{Rad: 1000, SrgbClr: &SrgbClr{Val: "FF0000"}},
		OuterShdw:  &OuterShdw{BlurRad: 57150},
		Reflection: &Reflection{StA: 50000, Sy: -100000},
		SoftEdge:   &SoftEdge{Rad: 1000},
	}
	got := marshal(t, e)

	glowIdx := strings.Index(got, "<a:glow")
	shdwIdx := strings.Index(got, "<a:outerShdw")
	reflIdx := strings.Index(got, "<a:reflection")
	softIdx := strings.Index(got, "<a:softEdge")
	if glowIdx == -1 || shdwIdx == -1 || reflIdx == -1 || softIdx == -1 {
		t.Fatalf("expected all four effects present, got %s", got)
	}
	if !(glowIdx < shdwIdx && shdwIdx < reflIdx && reflIdx < softIdx) {
		t.Errorf("expected glow < outerShdw < reflection < softEdge (CT_EffectList sequence), got %s", got)
	}
}

func TestOuterShdw_TriStateRotWithShapeMarshalsExplicitZero(t *testing.T) {
	notRotating := TriState(false)
	shdw := &OuterShdw{
		BlurRad:      57150,
		Dist:         19050,
		Dir:          5400000,
		Algn:         "ctr",
		RotWithShape: &notRotating,
		SrgbClr:      &SrgbClr{Val: "000000", Alpha: &Alpha{Val: 63000}},
	}
	got := marshal(t, shdw)

	// The regression this guards: OnOff can only ever emit "1" or omit the
	// attribute — it cannot represent an explicit false. TriState must.
	if !strings.Contains(got, `rotWithShape="0"`) {
		t.Errorf("expected explicit rotWithShape=\"0\", got %s", got)
	}
	if !strings.Contains(got, `<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">`) {
		t.Errorf("expected all outerShdw attrs, got %s", got)
	}
	if !strings.Contains(got, `<a:srgbClr val="000000"><a:alpha val="63000">`) {
		t.Errorf("expected nested color with alpha, got %s", got)
	}
}

func TestOuterShdw_NilRotWithShapeOmitsAttr(t *testing.T) {
	shdw := &OuterShdw{BlurRad: 1000}
	got := marshal(t, shdw)
	if strings.Contains(got, "rotWithShape") {
		t.Errorf("expected no rotWithShape attr when unset, got %s", got)
	}
}

func TestGlow_SchemeClrVariant(t *testing.T) {
	g := &Glow{Rad: 5000, SchemeClr: &SchemeClr{Val: "accent1"}}
	got := marshal(t, g)
	if !strings.Contains(got, `<a:glow rad="5000"><a:schemeClr val="accent1">`) {
		t.Errorf("expected glow with schemeClr, got %s", got)
	}
}

func TestSoftEdge_RadIsRequiredAttr(t *testing.T) {
	s := &SoftEdge{Rad: 12700}
	got := marshal(t, s)
	if !strings.Contains(got, `<a:softEdge rad="12700">`) {
		t.Errorf("expected softEdge with rad attr, got %s", got)
	}
}
