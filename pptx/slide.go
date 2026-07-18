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
	"github.com/mmonterroca/pptxgo/pkg/errors"
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
	cSld        *CSld  // the slide's own p:cSld, so Background can set its Bg field after construction
	spTree      *SpTree
	nextShapeID uint32
}

// Background sets the slide's own background to a solid color, overriding
// whatever its layout/master would otherwise supply.
func (s *Slide) Background(c drawingml.Color) *Slide {
	s.cSld.Bg = &Bg{BgPr: &BgPr{Fill: drawingml.NewSolidFillRGB(c)}}
	return s
}

// BackgroundScheme sets the slide's own background to a theme color,
// referenced by scheme slot (e.g. SchemeAccent1) rather than an explicit
// RGB value.
func (s *Slide) BackgroundScheme(scheme SchemeColor) *Slide {
	s.cSld.Bg = &Bg{BgPr: &BgPr{Fill: drawingml.NewSolidFillScheme(string(scheme))}}
	return s
}

// AddTextBox adds a text-box shape at the given position and size (x, y, w,
// h, all in EMUs — see the Inches/Points helpers) and returns a handle for
// adding paragraphs to it. It is addShape with a rect outline and the
// txBox marker set — the same p:sp any other autoshape uses.
func (s *Slide) AddTextBox(x, y, w, h int) *TextBox {
	return s.addShape(ShapeRect, x, y, w, h, true)
}

// AddShape adds an autoshape with the given preset geometry (see the Shape*
// constants, e.g. ShapeEllipse — any of ST_ShapeType's 187 names is
// accepted, not only those with a named constant) at the given position
// and size (x, y, w, h, all in EMUs — see the Inches/Points helpers), and
// returns a handle for adding text content and setting fill/border/
// rotation/flip. A name outside ST_ShapeType is recorded as an error on
// the presentation (returned by Save), since a:prstGeom/@prst with an
// unrecognized value is a file PowerPoint refuses to open.
func (s *Slide) AddShape(prst PresetGeometry, x, y, w, h int) *ShapeRef {
	if !IsValidPresetGeometry(prst) {
		s.pres.addErr(errors.InvalidArgument("AddShape", "prst", string(prst),
			"must be a valid ST_ShapeType preset geometry name (e.g. \"rect\", \"ellipse\")"))
	}
	return s.addShape(prst, x, y, w, h, false)
}

// addShape is the shared core of AddTextBox and AddShape: both place a p:sp
// with a text body, differing only in preset geometry and whether p:cNvSpPr
// carries the txBox marker (required for PowerPoint to treat a bare
// rectangle as a text container rather than an autoshape).
func (s *Slide) addShape(prst PresetGeometry, x, y, w, h int, isTextBox bool) *ShapeRef {
	id := s.nextShapeID
	s.nextShapeID++

	name := "Shape"
	cNvSpPr := &CNvSpPr{}
	if isTextBox {
		name = "TextBox"
		cNvSpPr.TxBox = true
	}

	body := &drawingml.TextBody{BodyPr: &drawingml.BodyPr{}, LstStyle: &drawingml.LstStyle{}}
	spPr := &SpPr{
		Xfrm: &drawingml.Xfrm{
			Off: &drawingml.Off{X: x, Y: y},
			Ext: &drawingml.Ext{Cx: w, Cy: h},
		},
		PrstGeom: &drawingml.PrstGeom{Prst: string(prst), AvLst: &drawingml.AvLst{}},
	}
	shape := &Shape{
		NvSpPr: &NvSpPr{
			CNvPr:   &CNvPr{ID: id, Name: fmt.Sprintf("%s %d", name, id)},
			CNvSpPr: cNvSpPr,
			NvPr:    &NvPr{},
		},
		SpPr:   spPr,
		TxBody: body,
	}
	s.spTree.Content = append(s.spTree.Content, shape)

	return &ShapeRef{pres: s.pres, slidePath: s.path, body: body, spPr: spPr}
}

// AddImage adds an image at (x, y), both in EMUs, auto-sized from the
// image's own pixel dimensions at 96 DPI. Format (PNG, JPEG, or GIF) is
// auto-detected. For exact sizing use AddImageWithSize.
func (s *Slide) AddImage(path string, x, y int) *PictureRef {
	data, err := os.ReadFile(path)
	if err != nil {
		s.pres.addErr(err)
		return s.addPicture(nil, x, y, 0, 0, false)
	}
	return s.addPicture(data, x, y, 0, 0, false)
}

// AddImageWithSize adds an image at (x, y) with an explicit size (w, h),
// all in EMUs — including a (0, 0) size, which is used as given rather
// than falling back to auto-sizing. Format is auto-detected; the image's
// own pixel dimensions are otherwise ignored in favor of w and h.
func (s *Slide) AddImageWithSize(path string, x, y, w, h int) *PictureRef {
	data, err := os.ReadFile(path)
	if err != nil {
		s.pres.addErr(err)
		return s.addPicture(nil, x, y, w, h, true)
	}
	return s.addPicture(data, x, y, w, h, true)
}

// AddImageFromBytes adds an in-memory image (PNG, JPEG, or GIF, format
// auto-detected) at (x, y), auto-sized from its pixel dimensions at 96 DPI.
// Empty data records an error on the presentation (returned by Save).
func (s *Slide) AddImageFromBytes(data []byte, x, y int) *PictureRef {
	if len(data) == 0 {
		s.pres.addErr(errors.InvalidArgument("AddImageFromBytes", "data", len(data), "image data is empty"))
		return s.addPicture(nil, x, y, 0, 0, false)
	}
	return s.addPicture(data, x, y, 0, 0, false)
}

// AddImageFromBytesWithSize adds an in-memory image at (x, y) with an
// explicit size (w, h), all in EMUs — including a (0, 0) size, which is
// used as given rather than falling back to auto-sizing. Empty data
// records an error on the presentation (returned by Save).
func (s *Slide) AddImageFromBytesWithSize(data []byte, x, y, w, h int) *PictureRef {
	if len(data) == 0 {
		s.pres.addErr(errors.InvalidArgument("AddImageFromBytesWithSize", "data", len(data), "image data is empty"))
		return s.addPicture(nil, x, y, w, h, true)
	}
	return s.addPicture(data, x, y, w, h, true)
}

// addPicture is the shared path for all four AddImage* variants.
// useExplicitSize distinguishes the WithSize variants (which use w, h
// exactly as given, even (0, 0)) from the auto-sizing ones (which ignore w
// and h and use the image's own pixel dimensions, scaled at 96 DPI) — a
// zero-valued w/h is not itself the auto-size signal, since a caller who
// explicitly asked for a (0, 0) picture should get one, not a silently
// different size. A nil data slice is the "inert picture" signal: every
// caller that passes nil must first accumulate an error on the
// presentation (the file-read paths on an os.ReadFile failure, the
// from-bytes paths on empty input), so Save refuses to write rather than
// emitting a p:pic with no blipFill. addPicture still returns a usable, if
// imageless, *PictureRef so a long fluent chain doesn't need a nil check
// after every call.
func (s *Slide) addPicture(data []byte, x, y, w, h int, useExplicitSize bool) *PictureRef {
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
		// prepareImage returns the bytes to actually embed (an EXIF-rotated
		// JPEG is re-encoded upright; every other image passes through
		// untouched) plus that image's true pixel dimensions.
		data, wPx, hPx, contentType, ext, err = prepareImage(data)
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

	if !useExplicitSize {
		w, h = wPx*emuPerPixel96DPI, hPx*emuPerPixel96DPI
	}
	spPr.Xfrm = &drawingml.Xfrm{Off: &drawingml.Off{X: x, Y: y}, Ext: &drawingml.Ext{Cx: w, Cy: h}}

	basename := s.pres.mediaBasename(data, contentType, ext)

	rid, relErr := s.pres.pkg.Relationships(s.path).AddImage("../media/" + basename)
	if relErr != nil {
		panic(relErr) // static, well-formed arguments; cannot fail
	}

	pic.BlipFill = NewBlipFill(rid)
	s.spTree.Content = append(s.spTree.Content, pic)

	return &PictureRef{pres: s.pres, spPr: spPr}
}
