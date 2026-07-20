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

func TestLn_MarshalsWidthAndSolidFill(t *testing.T) {
	ln := NewLn(Color{R: 0x1F, G: 0x49, B: 0x7D}, 12700)
	got := marshal(t, ln)

	if !strings.Contains(got, `<a:ln w="12700">`) {
		t.Errorf("expected a:ln with w attr, got %s", got)
	}
	if !strings.Contains(got, `<a:solidFill><a:srgbClr val="1F497D">`) {
		t.Errorf("expected nested solidFill, got %s", got)
	}
}

func TestLn_ZeroWidthOmitsAttr(t *testing.T) {
	ln := &Ln{}
	got := marshal(t, ln)
	if strings.Contains(got, `w=`) {
		t.Errorf("expected no w attr for zero width, got %s", got)
	}
}

func TestLn_PrstDashMarshalsAfterSolidFill(t *testing.T) {
	ln := NewLn(Color{R: 0x1F, G: 0x49, B: 0x7D}, 12700)
	ln.PrstDash = &PrstDash{Val: "dash"}
	got := marshal(t, ln)

	if !strings.Contains(got, `<a:prstDash val="dash">`) {
		t.Errorf("expected a:prstDash element, got %s", got)
	}
	fillIdx := strings.Index(got, "<a:solidFill>")
	dashIdx := strings.Index(got, "<a:prstDash")
	if fillIdx == -1 || dashIdx == -1 || fillIdx > dashIdx {
		t.Errorf("expected a:solidFill before a:prstDash (CT_LineProperties sequence), got %s", got)
	}
}

func TestLn_CapIsAttrAlongsideW(t *testing.T) {
	ln := &Ln{W: 12700, Cap: "rnd"}
	got := marshal(t, ln)

	if !strings.Contains(got, `<a:ln w="12700" cap="rnd">`) {
		t.Errorf("expected cap attr alongside w, got %s", got)
	}
}

func TestLn_JoinChoiceMiterCarriesLim(t *testing.T) {
	ln := &Ln{W: 12700, Miter: &LnMiter{Lim: 800000}}
	got := marshal(t, ln)

	if !strings.Contains(got, `<a:miter lim="800000">`) {
		t.Errorf("expected a:miter with lim attr, got %s", got)
	}
	if strings.Contains(got, "a:round") || strings.Contains(got, "a:bevel") {
		t.Errorf("expected only miter, no round/bevel, got %s", got)
	}
}

func TestLn_RoundAndBevelAreEmptyElements(t *testing.T) {
	round := marshal(t, &Ln{W: 1, Round: &LnRound{}})
	if !strings.Contains(round, "<a:round></a:round>") {
		t.Errorf("expected empty a:round element, got %s", round)
	}

	bevel := marshal(t, &Ln{W: 1, Bevel: &LnBevel{}})
	if !strings.Contains(bevel, "<a:bevel></a:bevel>") {
		t.Errorf("expected empty a:bevel element, got %s", bevel)
	}
}

func TestLn_HeadEndAndTailEndMarshalUnderDistinctTags(t *testing.T) {
	// LineEnd deliberately has no XMLName of its own (see its doc comment) —
	// this is the regression test for that: the SAME struct type must
	// marshal as a:headEnd or a:tailEnd purely from the field's own tag.
	ln := &Ln{
		W:       12700,
		HeadEnd: &LineEnd{Type: "triangle", W: "med", Len: "med"},
		TailEnd: &LineEnd{Type: "arrow", W: "lg", Len: "lg"},
	}
	got := marshal(t, ln)

	if !strings.Contains(got, `<a:headEnd type="triangle" w="med" len="med">`) {
		t.Errorf("expected a:headEnd with triangle type, got %s", got)
	}
	if !strings.Contains(got, `<a:tailEnd type="arrow" w="lg" len="lg">`) {
		t.Errorf("expected a:tailEnd with arrow type, got %s", got)
	}
	headIdx := strings.Index(got, "<a:headEnd")
	tailIdx := strings.Index(got, "<a:tailEnd")
	if headIdx == -1 || tailIdx == -1 || headIdx > tailIdx {
		t.Errorf("expected a:headEnd before a:tailEnd (CT_LineProperties sequence), got %s", got)
	}
}
