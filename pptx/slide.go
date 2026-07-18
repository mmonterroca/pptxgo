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

// AddTable adds a rows x cols table at the given position and overall size
// (x, y, w, h, all in EMUs — see the Inches/Points helpers), with column
// widths and row heights initially split evenly across w and h, and
// returns a handle for setting cell content and column/row sizing. Every
// cell starts with an empty txBody, the same "always at least one a:p"
// schema guarantee AddTextBox gives a fresh text box (drawingml.TextBody's
// own MarshalXML fills it in even without an explicit AddParagraph call).
//
// Unlike a shape or picture, a table is wrapped in a p:graphicFrame, not a
// p:sp — a:tbl content lives entirely inline in the slide's own XML, with
// no separate part or relationship the way an image needs one.
func (s *Slide) AddTable(rows, cols, x, y, w, h int) *Table {
	id := s.nextShapeID
	s.nextShapeID++

	colW := w / cols
	grid := &drawingml.TblGrid{}
	for c := 0; c < cols; c++ {
		grid.GridCol = append(grid.GridCol, &drawingml.GridCol{W: colW})
	}

	rowH := h / rows
	tbl := &drawingml.Tbl{
		TblPr: &drawingml.TblPr{
			FirstRow:     true,
			BandRow:      true,
			TableStyleID: &drawingml.TableStyleID{Value: drawingml.DefaultTableStyleID},
		},
		TblGrid: grid,
	}
	for r := 0; r < rows; r++ {
		tr := &drawingml.Tr{H: rowH}
		for c := 0; c < cols; c++ {
			tr.Tcs = append(tr.Tcs, &drawingml.Tc{
				TxBody: &drawingml.TextBody{BodyPr: &drawingml.BodyPr{}, LstStyle: &drawingml.LstStyle{}},
			})
		}
		tbl.Trs = append(tbl.Trs, tr)
	}

	frame := &GraphicFrame{
		NvGraphicFramePr: &NvGraphicFramePr{
			CNvPr:             &CNvPr{ID: id, Name: fmt.Sprintf("Table %d", id)},
			CNvGraphicFramePr: &CNvGraphicFramePr{},
			NvPr:              &NvPr{},
		},
		Xfrm: &GraphicFrameXfrm{
			Off: &drawingml.Off{X: x, Y: y},
			Ext: &drawingml.Ext{Cx: w, Cy: h},
		},
		Graphic: drawingml.NewGraphic(&drawingml.GraphicData{URI: drawingml.GraphicDataURITable, Inner: tbl}),
	}
	s.spTree.Content = append(s.spTree.Content, frame)

	return &Table{pres: s.pres, tbl: tbl, rows: rows, cols: cols}
}
