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
	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// Table is a handle onto a table placed via Slide.AddTable, for setting
// cell content and column/row sizing. ext is the enclosing p:graphicFrame's
// own extent (p:xfrm/a:ext), which ColumnWidth/RowHeight keep in sync with
// the table's actual total width/height — without that, resizing a column
// would change a:tblGrid but leave the frame's own bounding box stale.
// slidePath is the owning slide's part path, threaded down to each cell's
// Paragraph so Paragraph.Hyperlink (called from inside a table cell) scopes
// its relationship to the slide's own .rels, not the package root's — see
// ShapeRef.slidePath for the same requirement on shapes/text boxes.
type Table struct {
	pres      *Presentation
	slidePath string
	tbl       *drawingml.Tbl
	ext       *drawingml.Ext
}

// Cell returns a handle for setting the content of the cell at (row, col),
// both 0-indexed. Cell panics if row or col is out of range — the same
// contract a Go slice index gives, since a table's shape is fixed at
// AddTable and never grows.
func (t *Table) Cell(row, col int) *TableCell {
	tc := t.tbl.Trs[row].Tcs[col]
	return &TableCell{pres: t.pres, slidePath: t.slidePath, tc: tc}
}

// ColumnWidth sets the width of the given column, in EMUs (see the
// Inches/Points helpers), and recomputes the enclosing graphic frame's
// overall width (p:xfrm/a:ext/@cx) as the new sum of all column widths —
// AddTable splits the table's overall width evenly across columns to
// start, but the table's actual rendered width is always the sum of its
// column widths, and the frame's own extent must track that or the two
// disagree. An out-of-range col is recorded as an error on the
// presentation (returned by Save) and leaves the width unset, rather than
// panicking.
func (t *Table) ColumnWidth(col, widthEMU int) *Table {
	if col < 0 || col >= len(t.tbl.TblGrid.GridCol) {
		t.pres.addErr(errors.InvalidArgument("Table.ColumnWidth", "col", col, "out of range for this table's column count"))
		return t
	}
	t.tbl.TblGrid.GridCol[col].W = widthEMU

	total := 0
	for _, gc := range t.tbl.TblGrid.GridCol {
		total += gc.W
	}
	t.ext.Cx = total

	return t
}

// RowHeight sets the height of the given row, in EMUs, and recomputes the
// enclosing graphic frame's overall height (p:xfrm/a:ext/@cy) as the new
// sum of all row heights — see ColumnWidth for why the frame's extent
// must track the table's actual size. An out-of-range row is recorded as
// an error on the presentation (returned by Save) and leaves the height
// unset, rather than panicking.
func (t *Table) RowHeight(row, heightEMU int) *Table {
	if row < 0 || row >= len(t.tbl.Trs) {
		t.pres.addErr(errors.InvalidArgument("Table.RowHeight", "row", row, "out of range for this table's row count"))
		return t
	}
	t.tbl.Trs[row].H = heightEMU

	total := 0
	for _, tr := range t.tbl.Trs {
		total += tr.H
	}
	t.ext.Cy = total

	return t
}

// MergeCells merges the rectangular region of cells from (fromRow, fromCol)
// to (toRow, toCol), inclusive, both 0-indexed. The encoding follows the
// convention real PowerPoint-authored files use (extracted from a
// python-pptx-generated table and confirmed against the OpenXML SDK
// validator, not hand-derived from the schema alone) — a schema-valid but
// wrong encoding passes validation yet gets silently "repaired" by
// PowerPoint on open:
//
//   - The region never loses a <a:tc>: every cell stays in its row so each
//     row's cell count keeps matching a:tblGrid's column count, which
//     PowerPoint treats as corrupt if it disagrees.
//   - The anchor (fromRow, fromCol) carries GridSpan/RowSpan (only set when
//     greater than 1 — the schema's implicit default).
//   - A cell in the anchor's row but a later column also carries RowSpan
//     (it heads its own column's vertical span within the region) plus
//     HMerge.
//   - A cell in the anchor's column but a later row also carries GridSpan
//     (it heads its own row's horizontal span within the region) plus
//     VMerge.
//   - Every other (interior/corner) cell carries only HMerge and VMerge.
//
// An out-of-range or inverted (from > to) region, or one overlapping a cell
// already part of another merge, is recorded as an error on the
// presentation (returned by Save) and leaves the table unchanged.
//
// PowerPoint itself renders only the anchor cell's own content for a merged
// region — a non-anchor cell's txBody, even though the XML still carries
// it, never appears on the rendered slide. Populating a cell with
// Cell(...).Text(...) before merging it away is a natural call order (see
// examples/01_basic), so MergeCells transfers any non-anchor cell's
// existing paragraphs into the anchor (appended, in row-major order, after
// the anchor's own) rather than silently discarding them — the same
// mark-don't-delete principle the merge encoding itself follows.
func (t *Table) MergeCells(fromRow, fromCol, toRow, toCol int) *Table {
	if fromRow < 0 || fromCol < 0 || toRow >= len(t.tbl.Trs) || fromRow > toRow {
		t.pres.addErr(errors.InvalidArgument("Table.MergeCells", "fromRow/toRow", []int{fromRow, toRow}, "out of range or inverted for this table's row count"))
		return t
	}
	if toCol >= len(t.tbl.TblGrid.GridCol) || fromCol > toCol {
		t.pres.addErr(errors.InvalidArgument("Table.MergeCells", "fromCol/toCol", []int{fromCol, toCol}, "out of range or inverted for this table's column count"))
		return t
	}

	for r := fromRow; r <= toRow; r++ {
		for c := fromCol; c <= toCol; c++ {
			tc := t.tbl.Trs[r].Tcs[c]
			if tc.GridSpan > 1 || tc.RowSpan > 1 || tc.HMerge || tc.VMerge {
				t.pres.addErr(errors.InvalidArgument("Table.MergeCells", "region", []int{fromRow, fromCol, toRow, toCol}, "overlaps a cell already part of another merge"))
				return t
			}
		}
	}

	anchor := t.tbl.Trs[fromRow].Tcs[fromCol]
	for r := fromRow; r <= toRow; r++ {
		for c := fromCol; c <= toCol; c++ {
			if r == fromRow && c == fromCol {
				continue
			}
			tc := t.tbl.Trs[r].Tcs[c]
			if tc.TxBody == nil || len(tc.TxBody.Paragraphs) == 0 {
				continue
			}
			if anchor.TxBody == nil {
				anchor.TxBody = drawingml.NewTextBody()
			}
			anchor.TxBody.Paragraphs = append(anchor.TxBody.Paragraphs, tc.TxBody.Paragraphs...)
			tc.TxBody = drawingml.NewTextBody()
		}
	}

	rowSpan := toRow - fromRow + 1
	gridSpan := toCol - fromCol + 1

	for r := fromRow; r <= toRow; r++ {
		for c := fromCol; c <= toCol; c++ {
			tc := t.tbl.Trs[r].Tcs[c]
			switch {
			case r == fromRow && c == fromCol:
				if gridSpan > 1 {
					tc.GridSpan = gridSpan
				}
				if rowSpan > 1 {
					tc.RowSpan = rowSpan
				}
			case r == fromRow:
				if rowSpan > 1 {
					tc.RowSpan = rowSpan
				}
				tc.HMerge = true
			case c == fromCol:
				if gridSpan > 1 {
					tc.GridSpan = gridSpan
				}
				tc.VMerge = true
			default:
				tc.HMerge = true
				tc.VMerge = true
			}
		}
	}

	return t
}

// TableCell is a handle onto a single table cell (a:tc), returned by
// Table.Cell.
type TableCell struct {
	pres      *Presentation
	slidePath string
	tc        *drawingml.Tc
}

// AddParagraph appends a new, empty paragraph to the cell and returns a
// handle for adding runs and formatting to it — the same Paragraph type
// Slide.AddTextBox uses, so bold, alignment, hyperlinks, and every other
// text-formatting method apply equally inside a table cell.
func (c *TableCell) AddParagraph() *Paragraph {
	if c.tc.TxBody == nil {
		c.tc.TxBody = drawingml.NewTextBody()
	}
	p := &drawingml.Paragraph{}
	c.tc.TxBody.Paragraphs = append(c.tc.TxBody.Paragraphs, p)
	return &Paragraph{pres: c.pres, slidePath: c.slidePath, p: p}
}

// Text is shorthand for AddParagraph().Text(s) — the common case of a
// cell holding one plain line of text.
func (c *TableCell) Text(s string) *Paragraph {
	return c.AddParagraph().Text(s)
}
