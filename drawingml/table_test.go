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

func TestTbl_TblPrBeforeTblGridBeforeTr(t *testing.T) {
	tbl := &Tbl{
		TblPr:   &TblPr{FirstRow: true, TableStyleID: &TableStyleID{Value: DefaultTableStyleID}},
		TblGrid: &TblGrid{GridCol: []*GridCol{{W: 100}, {W: 200}}},
		Trs:     []*Tr{{H: 50, Tcs: []*Tc{{}, {}}}},
	}
	got := marshal(t, tbl)

	tblPrIdx := strings.Index(got, "<a:tblPr")
	tblGridIdx := strings.Index(got, "<a:tblGrid")
	trIdx := strings.Index(got, "<a:tr")
	if tblPrIdx == -1 || tblGridIdx == -1 || trIdx == -1 {
		t.Fatalf("expected a:tblPr, a:tblGrid, and a:tr all present, got %s", got)
	}
	if !(tblPrIdx < tblGridIdx && tblGridIdx < trIdx) {
		t.Errorf("expected a:tblPr < a:tblGrid < a:tr order, got %s", got)
	}
	if !strings.Contains(got, `firstRow="1"`) {
		t.Errorf("expected firstRow=\"1\", got %s", got)
	}
	if !strings.Contains(got, "<a:tableStyleId>"+DefaultTableStyleID+"</a:tableStyleId>") {
		t.Errorf("expected tableStyleId chardata, got %s", got)
	}
}

func TestTblGrid_GridColWidthsInOrder(t *testing.T) {
	grid := &TblGrid{GridCol: []*GridCol{{W: 100}, {W: 200}, {W: 300}}}
	got := marshal(t, grid)

	for _, want := range []string{`<a:gridCol w="100">`, `<a:gridCol w="200">`, `<a:gridCol w="300">`} {
		if !strings.Contains(got, want) {
			t.Errorf("expected %q in %s", want, got)
		}
	}
	firstIdx := strings.Index(got, `w="100"`)
	secondIdx := strings.Index(got, `w="200"`)
	thirdIdx := strings.Index(got, `w="300"`)
	if !(firstIdx < secondIdx && secondIdx < thirdIdx) {
		t.Errorf("expected columns in declared order, got %s", got)
	}
}

func TestTr_HAttrAndCellsInOrder(t *testing.T) {
	tr := &Tr{H: 457200, Tcs: []*Tc{
		{TxBody: &TextBody{Paragraphs: []*Paragraph{{Content: []any{&Run{Text: NewText("A")}}}}}},
		{TxBody: &TextBody{Paragraphs: []*Paragraph{{Content: []any{&Run{Text: NewText("B")}}}}}},
	}}
	got := marshal(t, tr)

	if !strings.Contains(got, `<a:tr h="457200">`) {
		t.Errorf("expected h=\"457200\" attr, got %s", got)
	}
	aIdx := strings.Index(got, ">A<")
	bIdx := strings.Index(got, ">B<")
	if aIdx == -1 || bIdx == -1 || aIdx > bIdx {
		t.Errorf("expected cell A before cell B, got %s", got)
	}
}

func TestTc_TxBodyBeforeMergeAttrsDoesNotMatterButPresent(t *testing.T) {
	tc := &Tc{
		TxBody:   &TextBody{Paragraphs: []*Paragraph{{Content: []any{&Run{Text: NewText("Header")}}}}},
		GridSpan: 2,
	}
	got := marshal(t, tc)

	if !strings.Contains(got, "<a:txBody>") {
		t.Errorf("expected a:txBody, got %s", got)
	}
	if !strings.Contains(got, `gridSpan="2"`) {
		t.Errorf("expected gridSpan=\"2\", got %s", got)
	}
}

func TestTc_UnmergedCellOmitsSpanAndMergeAttrs(t *testing.T) {
	tc := &Tc{TxBody: &TextBody{}}
	got := marshal(t, tc)

	for _, unwanted := range []string{"gridSpan", "rowSpan", "hMerge", "vMerge"} {
		if strings.Contains(got, unwanted) {
			t.Errorf("expected no %s attr on an unmerged cell, got %s", unwanted, got)
		}
	}
}

func TestTcPr_BorderSidesPrecedeFillGroup(t *testing.T) {
	// TcBorderLn (not Ln) is deliberately used here — see its own doc
	// comment for why Ln's fixed "a:ln" XMLName can't fill this role.
	tcPr := &TcPr{
		Anchor:    "ctr",
		LnT:       &TcBorderLn{W: 12700, SolidFill: NewSolidFillRGB(Color{R: 0x1F, G: 0x49, B: 0x7D})},
		LnBlToTr:  &TcBorderLn{W: 6350},
		SolidFill: NewSolidFillRGB(Color{R: 0xFF}),
	}
	got := marshal(t, tcPr)

	if !strings.Contains(got, `<a:lnT w="12700">`) {
		t.Errorf("expected a:lnT, got %s", got)
	}
	if !strings.Contains(got, `<a:lnBlToTr w="6350">`) {
		t.Errorf("expected a:lnBlToTr, got %s", got)
	}
	lnTIdx := strings.Index(got, "<a:lnT ")
	lnBlToTrIdx := strings.Index(got, "<a:lnBlToTr")
	// The cell's own fill (not lnT's border color) is the last solidFill,
	// identifiable by its distinct FF0000 value.
	fillIdx := strings.Index(got, `<a:srgbClr val="FF0000">`)
	if lnTIdx == -1 || lnBlToTrIdx == -1 || fillIdx == -1 || !(lnTIdx < lnBlToTrIdx && lnBlToTrIdx < fillIdx) {
		t.Errorf("expected lnT < lnBlToTr < solidFill (CT_TableCellProperties sequence), got %s", got)
	}
}

func TestTc_HMergeVMergeMarshalAsOneZeroNotTrueFalse(t *testing.T) {
	// Real PowerPoint output (and pptx.Table.MergeCells) always writes
	// hMerge="1"/vMerge="1", never Go's default bool "true"/"false" — both
	// are valid per xsd:boolean's lexical space, but matching the real
	// convention is safer for non-SDK consumers.
	tc := &Tc{TxBody: &TextBody{}, HMerge: true, VMerge: true}
	got := marshal(t, tc)

	if !strings.Contains(got, `hMerge="1"`) {
		t.Errorf("expected hMerge=\"1\", got %s", got)
	}
	if !strings.Contains(got, `vMerge="1"`) {
		t.Errorf("expected vMerge=\"1\", got %s", got)
	}
	if strings.Contains(got, "true") || strings.Contains(got, "false") {
		t.Errorf("expected no true/false lexical form, got %s", got)
	}
}
