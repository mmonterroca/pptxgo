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

import "encoding/xml"

// DefaultTableStyleID is Office's own built-in default table style
// ("Medium Style 2 - Accent 1"), the one PowerPoint itself applies to a
// freshly inserted table. Used by pptx.Slide.AddTable so a table renders
// with normal banded rows and header shading instead of bare, unstyled
// gridlines.
const DefaultTableStyleID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

// Tbl is a:tbl (CT_Table): a DrawingML table, the content a p:graphicFrame
// wraps when its GraphicData.URI is GraphicDataURITable. Field order
// mirrors the schema: TblPr, then the required TblGrid, then zero or more
// rows.
type Tbl struct {
	XMLName xml.Name `xml:"a:tbl"`
	TblPr   *TblPr   `xml:"a:tblPr,omitempty"`
	TblGrid *TblGrid `xml:"a:tblGrid"`
	Trs     []*Tr    `xml:"a:tr,omitempty"`
}

// TblPr is a:tblPr (CT_TableProperties): table-wide styling. FirstRow and
// BandRow mark the first row as a header and every other row as banded,
// respectively, for whichever style TableStyleID references.
type TblPr struct {
	XMLName      xml.Name      `xml:"a:tblPr"`
	FirstRow     OnOff         `xml:"firstRow,attr,omitempty"`
	BandRow      OnOff         `xml:"bandRow,attr,omitempty"`
	TableStyleID *TableStyleID `xml:"a:tableStyleId,omitempty"`
}

// TableStyleID is a:tableStyleId: a GUID reference into the theme's
// table-style list (see DefaultTableStyleID).
type TableStyleID struct {
	XMLName xml.Name `xml:"a:tableStyleId"`
	Value   string   `xml:",chardata"`
}

// TblGrid is a:tblGrid (CT_TableGrid): the table's column definitions.
type TblGrid struct {
	XMLName xml.Name   `xml:"a:tblGrid"`
	GridCol []*GridCol `xml:"a:gridCol,omitempty"`
}

// GridCol is a:gridCol (CT_TableCol): one column's width, in EMUs.
type GridCol struct {
	XMLName xml.Name `xml:"a:gridCol"`
	W       int      `xml:"w,attr"`
}

// Tr is a:tr (CT_TableRow): one table row — a height in EMUs and its
// cells, in left-to-right order.
type Tr struct {
	XMLName xml.Name `xml:"a:tr"`
	H       int      `xml:"h,attr"`
	Tcs     []*Tc    `xml:"a:tc,omitempty"`
}

// Tc is a:tc (CT_TableCell): one table cell. TxBody comes before the
// merge/span attributes in document order (they're attributes, so actual
// attribute order doesn't affect validity) but after it in the schema's
// own element sequence — TxBody is a:tc's only modeled child element.
// GridSpan/RowSpan greater than 1, and HMerge/VMerge, mark a cell as the
// anchor or continuation of a horizontal/vertical merge; their zero values
// (1, false, false) mean "not merged", the common case, and are the only
// case pptx.Slide.AddTable's builder currently produces.
type Tc struct {
	XMLName  xml.Name  `xml:"a:tc"`
	TxBody   *TextBody `xml:"a:txBody,omitempty"`
	GridSpan int       `xml:"gridSpan,attr,omitempty"`
	RowSpan  int       `xml:"rowSpan,attr,omitempty"`
	HMerge   OnOff     `xml:"hMerge,attr,omitempty"`
	VMerge   OnOff     `xml:"vMerge,attr,omitempty"`
}
