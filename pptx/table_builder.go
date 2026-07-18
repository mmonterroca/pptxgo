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

import "github.com/mmonterroca/pptxgo/drawingml"

// Table is a handle onto a table placed via Slide.AddTable, for setting
// cell content and column/row sizing.
type Table struct {
	pres *Presentation
	tbl  *drawingml.Tbl
	rows int
	cols int
}

// Cell returns a handle for setting the content of the cell at (row, col),
// both 0-indexed. Cell panics if row or col is out of range — the same
// contract a Go slice index gives, since a table's shape is fixed at
// AddTable and never grows.
func (t *Table) Cell(row, col int) *TableCell {
	tc := t.tbl.Trs[row].Tcs[col]
	return &TableCell{pres: t.pres, tc: tc}
}

// ColumnWidth sets the width of the given column, in EMUs (see the
// Inches/Points helpers). AddTable splits the table's overall width
// evenly across columns to start; the table's actual rendered width is
// the sum of its column widths, so changing one changes the total.
func (t *Table) ColumnWidth(col, widthEMU int) *Table {
	t.tbl.TblGrid.GridCol[col].W = widthEMU
	return t
}

// RowHeight sets the height of the given row, in EMUs.
func (t *Table) RowHeight(row, heightEMU int) *Table {
	t.tbl.Trs[row].H = heightEMU
	return t
}

// TableCell is a handle onto a single table cell (a:tc), returned by
// Table.Cell.
type TableCell struct {
	pres *Presentation
	tc   *drawingml.Tc
}

// AddParagraph appends a new, empty paragraph to the cell and returns a
// handle for adding runs and formatting to it — the same Paragraph type
// Slide.AddTextBox uses, so bold, alignment, and every other text-
// formatting method apply equally inside a table cell.
func (c *TableCell) AddParagraph() *Paragraph {
	if c.tc.TxBody == nil {
		c.tc.TxBody = &drawingml.TextBody{BodyPr: &drawingml.BodyPr{}, LstStyle: &drawingml.LstStyle{}}
	}
	p := &drawingml.Paragraph{}
	c.tc.TxBody.Paragraphs = append(c.tc.TxBody.Paragraphs, p)
	return &Paragraph{pres: c.pres, p: p}
}

// Text is shorthand for AddParagraph().Text(s) — the common case of a
// cell holding one plain line of text.
func (c *TableCell) Text(s string) *Paragraph {
	return c.AddParagraph().Text(s)
}
