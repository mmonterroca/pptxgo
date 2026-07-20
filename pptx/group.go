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

	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// GroupShape is p:grpSp (CT_GroupShape): a set of shapes that move, resize,
// and rotate together in PowerPoint's authoring UI. Structurally parallel
// to SpTree (nvGrpSpPr + grpSpPr + an ordered Content) but its own type,
// not a reuse of SpTree — SpTree's own fixed "p:spTree" XMLName can't be
// renamed to "p:grpSp" via a field tag (the same conflict this package's
// other new-type-per-tag types already document). Content is an
// order-preserving `[]any`, the same pattern SpTree.Content uses:
// CT_GroupShape's own EG_ShapeElements group structurally allows shapes,
// pictures, tables, nested groups, and connectors interleaved in any order,
// and there is no custom MarshalXML to reorder them. The Group handle
// currently only routes AddShape/AddTextBox (p:sp) into it; nesting a
// picture, table, or sub-group is a possible follow-up (each needs its own
// real-PowerPoint check that the element actually moves with the group).
type GroupShape struct {
	XMLName   xml.Name   `xml:"p:grpSp"`
	NvGrpSpPr *NvGrpSpPr `xml:"p:nvGrpSpPr"`
	GrpSpPr   *GrpSpPr   `xml:"p:grpSpPr"`
	Content   []any      `xml:",any"`
}

// Group is a handle onto a p:grpSp, returned by Slide.AddGroup. AddShape
// and AddTextBox mirror Slide's own methods of the same name, but append
// into the group's own Content instead of the slide's top-level shape
// tree — a member shape's (x, y) is still a slide-absolute EMU position
// (see AddGroup's own doc comment on the group's 1:1 child coordinate
// space), so no coordinate translation happens here.
type Group struct {
	pres      *Presentation
	slidePath string
	slide     *Slide // shape IDs are slide-global, not per-group — see Slide.allocID
	grp       *GroupShape
}

// AddTextBox adds a text-box shape to the group at the given position and
// size (x, y, w, h, all in EMUs, slide-absolute) and returns a handle for
// adding paragraphs to it — Slide.AddTextBox's counterpart for a group
// member.
func (g *Group) AddTextBox(x, y, w, h int) *TextBox {
	return g.addShape(ShapeRect, x, y, w, h, true)
}

// AddShape adds an autoshape to the group with the given preset geometry at
// the given position and size (x, y, w, h, all in EMUs, slide-absolute) —
// Slide.AddShape's counterpart for a group member. An invalid preset name
// is recorded as an error on the presentation (returned by Save), the same
// contract Slide.AddShape gives.
func (g *Group) AddShape(prst PresetGeometry, x, y, w, h int) *ShapeRef {
	if !IsValidPresetGeometry(prst) {
		g.pres.addErr(errors.InvalidArgument("Group.AddShape", "prst", string(prst),
			"must be a valid ST_ShapeType preset geometry name (e.g. \"rect\", \"ellipse\")"))
	}
	return g.addShape(prst, x, y, w, h, false)
}

// addShape is the shared core of Group.AddTextBox and Group.AddShape,
// mirroring Slide.addShape but appending into the group's own Content and
// drawing its shape id from the owning slide's slide-global allocator.
func (g *Group) addShape(prst PresetGeometry, x, y, w, h int, isTextBox bool) *ShapeRef {
	shape, ref := buildShape(g.pres, g.slidePath, g.slide.allocID(), prst, x, y, w, h, isTextBox)
	g.grp.Content = append(g.grp.Content, shape)
	return ref
}
