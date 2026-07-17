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
	"os"

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
	path        string // this slide's own part path, e.g. "ppt/slides/slide1.xml" — needed to own its media relationships
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
	spPr := &SpPr{
		Xfrm: &drawingml.Xfrm{
			Off: &drawingml.Off{X: x, Y: y},
			Ext: &drawingml.Ext{Cx: w, Cy: h},
		},
		PrstGeom: &drawingml.PrstGeom{Prst: "rect", AvLst: &drawingml.AvLst{}},
	}
	shape := &Shape{
		NvSpPr: &NvSpPr{
			CNvPr:   &CNvPr{ID: id, Name: fmt.Sprintf("TextBox %d", id)},
			CNvSpPr: &CNvSpPr{TxBox: true},
			NvPr:    &NvPr{},
		},
		SpPr:   spPr,
		TxBody: body,
	}
	s.spTree.Content = append(s.spTree.Content, shape)

	return &TextBox{pres: s.pres, body: body, spPr: spPr}
}

// AddImage adds an image at (x, y), both in EMUs, auto-sized from the
// image's own pixel dimensions at 96 DPI. Format (PNG, JPEG, or GIF) is
// auto-detected. For exact sizing use AddImageWithSize.
func (s *Slide) AddImage(path string, x, y int) *PictureRef {
	data, err := os.ReadFile(path)
	if err != nil {
		s.pres.addErr(err)
		return s.addPicture(nil, x, y, 0, 0)
	}
	return s.addPicture(data, x, y, 0, 0)
}

// AddImageWithSize adds an image at (x, y) with an explicit size (w, h),
// all in EMUs. Format is auto-detected; the image's own pixel dimensions
// are ignored in favor of w and h.
func (s *Slide) AddImageWithSize(path string, x, y, w, h int) *PictureRef {
	data, err := os.ReadFile(path)
	if err != nil {
		s.pres.addErr(err)
		return s.addPicture(nil, x, y, w, h)
	}
	return s.addPicture(data, x, y, w, h)
}

// AddImageFromBytes adds an in-memory image (PNG, JPEG, or GIF, format
// auto-detected) at (x, y), auto-sized from its pixel dimensions at 96 DPI.
func (s *Slide) AddImageFromBytes(data []byte, x, y int) *PictureRef {
	return s.addPicture(data, x, y, 0, 0)
}

// AddImageFromBytesWithSize adds an in-memory image at (x, y) with an
// explicit size (w, h), all in EMUs.
func (s *Slide) AddImageFromBytesWithSize(data []byte, x, y, w, h int) *PictureRef {
	return s.addPicture(data, x, y, w, h)
}

// addPicture is the shared path for all four AddImage* variants. w and h,
// when both non-zero, are used as-is (the WithSize variants); when both
// zero, the image's own pixel dimensions (scaled at 96 DPI) are used
// instead. When data is nil, a file read already failed and the error was
// accumulated by the caller — addPicture still returns a usable, if
// imageless, *PictureRef so a long fluent chain doesn't need a nil check
// after every call.
func (s *Slide) addPicture(data []byte, x, y, w, h int) *PictureRef {
	id := s.nextShapeID
	s.nextShapeID++

	spPr := &SpPr{PrstGeom: &drawingml.PrstGeom{Prst: "rect", AvLst: &drawingml.AvLst{}}}
	pic := &Picture{
		NvPicPr: &NvPicPr{
			CNvPr:    &CNvPr{ID: id, Name: fmt.Sprintf("Picture %d", id)},
			CNvPicPr: &CNvPicPr{PicLocks: &drawingml.PicLocks{NoChangeAspect: true}},
			NvPr:     &NvPr{},
		},
		SpPr: spPr,
	}

	var wPx, hPx int
	var contentType, ext string
	if data != nil {
		var err error
		wPx, hPx, contentType, ext, err = imageMeta(data)
		if err != nil {
			s.pres.addErr(err)
			data = nil // fall through to the inert path below
		}
	}

	if data == nil {
		spPr.Xfrm = &drawingml.Xfrm{Off: &drawingml.Off{X: x, Y: y}, Ext: &drawingml.Ext{Cx: w, Cy: h}}
		s.spTree.Content = append(s.spTree.Content, pic)
		return &PictureRef{pres: s.pres, spPr: spPr}
	}

	if w == 0 && h == 0 {
		w, h = wPx*emuPerPixel96DPI, hPx*emuPerPixel96DPI
	}
	spPr.Xfrm = &drawingml.Xfrm{Off: &drawingml.Off{X: x, Y: y}, Ext: &drawingml.Ext{Cx: w, Cy: h}}

	name := s.pres.pkg.IDs().NextID("image")
	part := "ppt/media/" + name + "." + ext
	s.pres.pkg.AddMediaPart(part, contentType, data)

	rid, relErr := s.pres.pkg.Relationships(s.path).AddImage("../media/" + name + "." + ext)
	if relErr != nil {
		panic(relErr) // static, well-formed arguments; cannot fail
	}

	pic.BlipFill = NewBlipFill(rid)
	s.spTree.Content = append(s.spTree.Content, pic)

	return &PictureRef{pres: s.pres, spPr: spPr}
}
