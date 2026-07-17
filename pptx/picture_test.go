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
	"encoding/xml"
	"strings"
	"testing"

	"github.com/mmonterroca/pptxgo/drawingml"
)

func marshal(t *testing.T, v any) string {
	t.Helper()
	b, err := xml.Marshal(v)
	if err != nil {
		t.Fatalf("xml.Marshal: %v", err)
	}
	return string(b)
}

func TestPicture_MarshalsSchemaOrderedElements(t *testing.T) {
	pic := &Picture{
		NvPicPr: &NvPicPr{
			CNvPr:    &CNvPr{ID: 2, Name: "Picture 2"},
			CNvPicPr: &CNvPicPr{PicLocks: &drawingml.PicLocks{NoChangeAspect: true}},
			NvPr:     &NvPr{},
		},
		BlipFill: NewBlipFill("rId2"),
		SpPr: &SpPr{
			Xfrm:     &drawingml.Xfrm{Off: &drawingml.Off{X: 100, Y: 200}, Ext: &drawingml.Ext{Cx: 300, Cy: 400}},
			PrstGeom: &drawingml.PrstGeom{Prst: "rect", AvLst: &drawingml.AvLst{}},
		},
	}
	got := marshal(t, pic)

	for _, want := range []string{
		"<p:pic>",
		"<p:nvPicPr>",
		"<p:blipFill>",
		`r:embed="rId2"`,
		"<p:spPr>",
		`noChangeAspect="true"`,
	} {
		if !strings.Contains(got, want) {
			t.Errorf("expected %q in %s", want, got)
		}
	}

	blipFillIdx := strings.Index(got, "<p:blipFill>")
	spPrIdx := strings.Index(got, "<p:spPr>")
	if blipFillIdx == -1 || spPrIdx == -1 || blipFillIdx > spPrIdx {
		t.Errorf("expected p:blipFill before p:spPr, got %s", got)
	}

	blipIdx := strings.Index(got, "<a:blip")
	stretchIdx := strings.Index(got, "<a:stretch")
	if blipIdx == -1 || stretchIdx == -1 || blipIdx > stretchIdx {
		t.Errorf("expected a:blip before a:stretch, got %s", got)
	}
}

func TestPicture_LnAfterFillInSpPr(t *testing.T) {
	spPr := &SpPr{
		Xfrm:     &drawingml.Xfrm{Ext: &drawingml.Ext{Cx: 1, Cy: 1}},
		PrstGeom: &drawingml.PrstGeom{Prst: "rect"},
		Ln:       drawingml.NewLn(drawingml.Color{R: 0, G: 0, B: 0}, 12700),
	}
	got := marshal(t, spPr)

	prstGeomIdx := strings.Index(got, "<a:prstGeom")
	lnIdx := strings.Index(got, "<a:ln")
	if prstGeomIdx == -1 || lnIdx == -1 || prstGeomIdx > lnIdx {
		t.Errorf("expected a:prstGeom before a:ln, got %s", got)
	}
}
