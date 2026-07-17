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

func TestRPr_CentipointsAndOnOffAttrs(t *testing.T) {
	rpr := &RPr{Sz: 3200, B: true, U: "sng"}
	got := marshal(t, rpr)

	for _, want := range []string{`sz="3200"`, `b="1"`, `u="sng"`} {
		if !strings.Contains(got, want) {
			t.Errorf("expected %q in %s", want, got)
		}
	}
}

func TestRPr_FalseOnOffAttrIsOmitted(t *testing.T) {
	rpr := &RPr{Sz: 1800, I: false}
	got := marshal(t, rpr)
	if strings.Contains(got, `i=`) {
		t.Errorf("expected i attr omitted when false, got %s", got)
	}
}

func TestRPr_SolidFillBeforeLatin(t *testing.T) {
	rpr := &RPr{
		SolidFill: NewSolidFillRGB(Color{R: 0x1F, G: 0x49, B: 0x7D}),
		Latin:     &Latin{Typeface: "Calibri"},
	}
	got := marshal(t, rpr)

	fillIdx := strings.Index(got, "<a:solidFill")
	latinIdx := strings.Index(got, "<a:latin")
	if fillIdx == -1 || latinIdx == -1 {
		t.Fatalf("expected both a:solidFill and a:latin in %s", got)
	}
	if fillIdx > latinIdx {
		t.Errorf("expected a:solidFill before a:latin, got %s", got)
	}
}

func TestRun_TextCarriesSpacePreserve(t *testing.T) {
	r := &Run{Text: NewText("  hi  ")}
	got := marshal(t, r)
	if !strings.Contains(got, `<a:t xml:space="preserve">  hi  </a:t>`) {
		t.Errorf("expected preserved-space a:t, got %s", got)
	}
}

func TestParagraph_PPrBeforeRuns(t *testing.T) {
	p := &Paragraph{
		PPr:  &PPr{Algn: "ctr"},
		Runs: []*Run{{Text: NewText("hello")}},
	}
	got := marshal(t, p)

	pprIdx := strings.Index(got, "<a:pPr")
	runIdx := strings.Index(got, "<a:r>")
	if pprIdx == -1 || runIdx == -1 {
		t.Fatalf("expected both a:pPr and a:r in %s", got)
	}
	if pprIdx > runIdx {
		t.Errorf("expected a:pPr before a:r, got %s", got)
	}
	if !strings.Contains(got, `algn="ctr"`) {
		t.Errorf("expected algn attr, got %s", got)
	}
}

// txBodyHost simulates how a host package (pptx.Shape) embeds TextBody: the
// element name comes from the field tag, not from TextBody itself.
type txBodyHost struct {
	XMLName xml.Name  `xml:"host"`
	TxBody  *TextBody `xml:"p:txBody"`
}

func TestTextBody_NameComesFromHostFieldTag(t *testing.T) {
	host := &txBodyHost{TxBody: &TextBody{
		BodyPr:   &BodyPr{},
		LstStyle: &LstStyle{},
		Paragraphs: []*Paragraph{
			{Runs: []*Run{{Text: NewText("Quarterly Results")}}},
		},
	}}
	got := marshal(t, host)

	if !strings.Contains(got, "<p:txBody>") || !strings.Contains(got, "</p:txBody>") {
		t.Errorf("expected p:txBody wrapper (name from field tag), got %s", got)
	}
	bodyPrIdx := strings.Index(got, "<a:bodyPr")
	lstStyleIdx := strings.Index(got, "<a:lstStyle")
	pIdx := strings.Index(got, "<a:p>")
	if bodyPrIdx == -1 || lstStyleIdx == -1 || pIdx == -1 {
		t.Fatalf("expected a:bodyPr, a:lstStyle, a:p all present, got %s", got)
	}
	if !(bodyPrIdx < lstStyleIdx && lstStyleIdx < pIdx) {
		t.Errorf("expected bodyPr < lstStyle < p order, got %s", got)
	}
}

func TestTextBody_EmitsOneParagraphWhenCallerAddedNone(t *testing.T) {
	host := &txBodyHost{TxBody: &TextBody{BodyPr: &BodyPr{}, LstStyle: &LstStyle{}}}
	got := marshal(t, host)

	if strings.Count(got, "<a:p>") != 1 {
		t.Errorf("expected exactly one empty <a:p> fallback, got %s", got)
	}
}
