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
