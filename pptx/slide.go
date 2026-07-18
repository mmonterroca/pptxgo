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
	"fmt"

	"github.com/mmonterroca/pptxgo/drawingml"
)

// firstShapeID is the first ID available for a slide's own shapes. ID 1 is
// permanently reserved for the shape tree's own nvGrpSpPr (see
// NewEmptySpTree), so every slide's shapes start numbering at 2.
const firstShapeID = 2

// Slide is a handle onto one already-registered slide part: the entry point
// for adding shapes to it. Obtain one via Presentation.AddSlide; the zero
// value is not usable.
type Slide struct {
	pres        *Presentation
	spTree      *SpTree
	nextShapeID uint32
}

// AddTextBox adds a text-box shape at the given position and size (x, y, w,
// h, all in EMUs — see the Inches/Points helpers) and returns a handle for
// adding paragraphs to it.
func (s *Slide) AddTextBox(x, y, w, h int) *TextBox {
	id := s.nextShapeID
	s.nextShapeID++

	body := &drawingml.TextBody{BodyPr: &drawingml.BodyPr{}, LstStyle: &drawingml.LstStyle{}}
	shape := &Shape{
		NvSpPr: &NvSpPr{
			CNvPr:   &CNvPr{ID: id, Name: fmt.Sprintf("TextBox %d", id)},
			CNvSpPr: &CNvSpPr{TxBox: true},
			NvPr:    &NvPr{},
		},
		SpPr: &SpPr{
			Xfrm: &drawingml.Xfrm{
				Off: &drawingml.Off{X: x, Y: y},
				Ext: &drawingml.Ext{Cx: w, Cy: h},
			},
			PrstGeom: &drawingml.PrstGeom{Prst: "rect", AvLst: &drawingml.AvLst{}},
		},
		TxBody: body,
	}
	s.spTree.Content = append(s.spTree.Content, shape)

	return &TextBox{pres: s.pres, body: body}
}
