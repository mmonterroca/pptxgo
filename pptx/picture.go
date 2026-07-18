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

// Picture is p:pic (CT_Picture): an image placed directly on a slide.
// Field order mirrors the schema: nvPicPr -> blipFill -> spPr.
type Picture struct {
	XMLName  xml.Name  `xml:"p:pic"`
	NvPicPr  *NvPicPr  `xml:"p:nvPicPr"`
	BlipFill *BlipFill `xml:"p:blipFill"`
	SpPr     *SpPr     `xml:"p:spPr"`
}

// NvPicPr is p:nvPicPr (CT_PictureNonVisual): the picture's non-visual
// properties.
type NvPicPr struct {
	XMLName  xml.Name  `xml:"p:nvPicPr"`
	CNvPr    *CNvPr    `xml:"p:cNvPr"`
	CNvPicPr *CNvPicPr `xml:"p:cNvPicPr"`
	NvPr     *NvPr     `xml:"p:nvPr"`
}

// CNvPicPr is p:cNvPicPr (CT_NonVisualPictureProperties): picture-specific
// non-visual drawing properties. PicLocks is set to noChangeAspect="1" by
// every builder path so a placed image can't be stretched out of
// proportion in the authoring UI.
type CNvPicPr struct {
	XMLName  xml.Name            `xml:"p:cNvPicPr"`
	PicLocks *drawingml.PicLocks `xml:"a:picLocks,omitempty"`
}

// BlipFill is p:blipFill (CT_BlipFillProperties): the image reference and
// how it fills the picture's frame. In PresentationML the wrapper element
// is p:blipFill (not a:blipFill) even though both children stay in the "a:"
// namespace — the same host-names-the-element pattern as p:txBody wrapping
// drawingml.TextBody in Fase 2.
type BlipFill struct {
	XMLName xml.Name           `xml:"p:blipFill"`
	Blip    *drawingml.Blip    `xml:"a:blip"`
	Stretch *drawingml.Stretch `xml:"a:stretch,omitempty"`
}

// NewBlipFill returns a BlipFill referencing relID and stretched to fill
// its shape's whole frame — the simplest, and by far most common, image
// fill mode.
func NewBlipFill(relID string) *BlipFill {
	return &BlipFill{
		Blip:    drawingml.NewBlip(relID),
		Stretch: &drawingml.Stretch{FillRect: &drawingml.FillRect{}},
	}
}
