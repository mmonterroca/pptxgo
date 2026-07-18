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

	"github.com/mmonterroca/pptxgo/drawingml"
)

// GraphicFrame is p:graphicFrame (CT_GraphicalObjectFrame): a slide shape
// that wraps non-p:sp, non-p:pic content — today, only a table (see
// Slide.AddTable); charts and embedded OLE objects use this same wrapper
// in later phases. Field order mirrors the schema:
// nvGraphicFramePr -> xfrm -> graphic.
type GraphicFrame struct {
	XMLName          xml.Name           `xml:"p:graphicFrame"`
	NvGraphicFramePr *NvGraphicFramePr  `xml:"p:nvGraphicFramePr"`
	Xfrm             *GraphicFrameXfrm  `xml:"p:xfrm"`
	Graphic          *drawingml.Graphic `xml:"a:graphic"`
}

// NvGraphicFramePr is p:nvGraphicFramePr (CT_GraphicalObjectFrameNonVisual):
// the graphic frame's non-visual properties — the same cNvPr/nvPr shape
// every other shape-like element carries (see NvSpPr, NvPicPr).
type NvGraphicFramePr struct {
	XMLName           xml.Name           `xml:"p:nvGraphicFramePr"`
	CNvPr             *CNvPr             `xml:"p:cNvPr"`
	CNvGraphicFramePr *CNvGraphicFramePr `xml:"p:cNvGraphicFramePr"`
	NvPr              *NvPr              `xml:"p:nvPr"`
}

// CNvGraphicFramePr is p:cNvGraphicFramePr (CT_NonVisualGraphicFrameProperties):
// non-visual drawing properties specific to a graphic frame. GraphicFrameLocks
// is optional and unset by Slide.AddTable — nothing yet needs to restrict
// resize/move on a table.
type CNvGraphicFramePr struct {
	XMLName           xml.Name                     `xml:"p:cNvGraphicFramePr"`
	GraphicFrameLocks *drawingml.GraphicFrameLocks `xml:"a:graphicFrameLocks,omitempty"`
}

// GraphicFrameXfrm is p:xfrm: a graphic frame's position and size, in
// EMUs. It exists as its own type, rather than reusing drawingml.Xfrm,
// because Xfrm's own fixed XMLName ("a:xfrm") always wins over any field
// tag that tries to rename it — encoding/xml's rule, already documented on
// drawingml.TextBody. p:xfrm has the identical content model
// (a:CT_Transform2D); only the element's own namespace prefix differs,
// because a graphic frame is a PresentationML (p:) element while a shape's
// a:xfrm is reused verbatim across formats.
type GraphicFrameXfrm struct {
	XMLName xml.Name       `xml:"p:xfrm"`
	Off     *drawingml.Off `xml:"a:off,omitempty"`
	Ext     *drawingml.Ext `xml:"a:ext"`
}
