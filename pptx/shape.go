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

// Shape is p:sp (CT_Shape): a single shape on a slide. Fase 2 only ever
// builds the text-box flavor (a rectangle with a txBody), but the wrapper
// models the full shape schema so pictures and placeholders can reuse it in
// later phases.
type Shape struct {
	XMLName xml.Name            `xml:"p:sp"`
	NvSpPr  *NvSpPr             `xml:"p:nvSpPr"`
	SpPr    *SpPr               `xml:"p:spPr"`
	TxBody  *drawingml.TextBody `xml:"p:txBody,omitempty"`
}

// NvSpPr is p:nvSpPr (CT_ShapeNonVisual): the shape's non-visual properties.
type NvSpPr struct {
	XMLName xml.Name `xml:"p:nvSpPr"`
	CNvPr   *CNvPr   `xml:"p:cNvPr"`
	CNvSpPr *CNvSpPr `xml:"p:cNvSpPr"`
	NvPr    *NvPr    `xml:"p:nvPr"`
}

// SpPr is p:spPr (CT_ShapeProperties): the shape's geometry and visual
// properties. A free text box needs an explicit a:xfrm — without one it has
// no position on the slide. Field order mirrors the schema:
// xfrm -> prstGeom -> (fill group) -> ln -> effectLst. This same struct is
// reused as-is for a p:pic's p:spPr (Fill, Gradient, and NoFill all stay nil
// there — a picture's fill is its blipFill, not a:solidFill/a:gradFill/
// a:noFill). Fill, Gradient, and NoFill are the schema's EG_FillProperties
// choice: at most one should ever be set — the ShapeRef builder methods
// (Fill, FillScheme, GradientFill, NoFill) enforce that by clearing the
// others whenever one is set.
type SpPr struct {
	XMLName   xml.Name             `xml:"p:spPr"`
	Xfrm      *drawingml.Xfrm      `xml:"a:xfrm,omitempty"`
	PrstGeom  *drawingml.PrstGeom  `xml:"a:prstGeom,omitempty"`
	Fill      *drawingml.SolidFill `xml:"a:solidFill,omitempty"`
	Gradient  *drawingml.GradFill  `xml:"a:gradFill,omitempty"`
	NoFill    *drawingml.NoFill    `xml:"a:noFill,omitempty"`
	Ln        *drawingml.Ln        `xml:"a:ln,omitempty"`
	EffectLst *drawingml.EffectLst `xml:"a:effectLst,omitempty"`
}
