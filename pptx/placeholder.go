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

// masterMarginPct, masterTitleHeightPct, masterTitleGapPct, and
// masterBottomGapPct lay out the master's own title and body placeholders
// as percentages of the slide canvas, so the geometry scales with whatever
// size WithSlideSize/WithStandard4x3 chose rather than assuming 16:9.
// These roughly mirror PowerPoint's own default "Title and Content" master
// layout proportions; nothing about placeholder inheritance depends on the
// exact values.
const (
	masterMarginPct      = 5  // left/right margin, and title's top margin
	masterTitleHeightPct = 15 // title placeholder height
	masterTitleGapPct    = 3  // gap between title and body placeholders
	masterBottomGapPct   = 5  // body placeholder's bottom margin
)

// newMasterSpTree builds the slide master's own shape tree: a title
// placeholder (type="title", the schema's default idx=0) and a body
// placeholder (type="body", idx=1) positioned proportionally within the
// given slide canvas size. These are the master's OWN placeholders — the
// ones every layout's and slide's same-typed/same-idx placeholder
// ultimately inherits position and formatting from when it sets none of
// its own.
func newMasterSpTree(slideWidthEMU, slideHeightEMU int) *SpTree {
	margin := slideWidthEMU * masterMarginPct / 100
	titleX, titleY := margin, slideHeightEMU*masterMarginPct/100
	titleW, titleH := slideWidthEMU-2*margin, slideHeightEMU*masterTitleHeightPct/100

	bodyX, bodyY := titleX, titleY+titleH+slideHeightEMU*masterTitleGapPct/100
	bodyW := titleW
	bodyH := slideHeightEMU - bodyY - slideHeightEMU*masterBottomGapPct/100

	spTree := NewEmptySpTree()
	spTree.Content = append(spTree.Content,
		newPlaceholderShape(2, "Title Placeholder 2", PlaceholderTitle, 0, &drawingml.Xfrm{
			Off: &drawingml.Off{X: titleX, Y: titleY},
			Ext: &drawingml.Ext{Cx: titleW, Cy: titleH},
		}),
		newPlaceholderShape(3, "Body Placeholder 3", PlaceholderBody, 1, &drawingml.Xfrm{
			Off: &drawingml.Off{X: bodyX, Y: bodyY},
			Ext: &drawingml.Ext{Cx: bodyW, Cy: bodyH},
		}),
	)
	return spTree
}

// newPlaceholderShape builds a p:sp carrying a p:ph placeholder marker
// instead of any preset geometry — a placeholder doesn't declare its own
// a:prstGeom; its outline comes from whichever same-typed, same-idx
// placeholder it inherits from. xfrm may be nil: a placeholder that omits
// a:xfrm entirely inherits position and size from its layout's (or, for a
// layout placeholder, its master's) matching placeholder. Shared by the
// master's own built-in placeholders (New) and, in later phases, standard
// layouts and Slide.AddPlaceholder.
func newPlaceholderShape(id uint32, name string, phType PlaceholderType, idx int, xfrm *drawingml.Xfrm) *Shape {
	return &Shape{
		NvSpPr: &NvSpPr{
			CNvPr:   &CNvPr{ID: id, Name: name},
			CNvSpPr: &CNvSpPr{},
			NvPr:    &NvPr{Ph: &Ph{Type: phType, Idx: idx}},
		},
		SpPr:   &SpPr{Xfrm: xfrm},
		TxBody: drawingml.NewTextBody(),
	}
}
