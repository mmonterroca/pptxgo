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

// PictureRef is a handle onto a placed image (a p:pic), returned by
// Slide.AddImage and its variants. Border is its only formatting method —
// an image has no text and no separate fill (its "fill" is the image
// itself, via blipFill).
type PictureRef struct {
	pres *Presentation
	spPr *SpPr
}

// Border sets the picture's outline to a solid color at the given width,
// in points (e.g. 0.75, 1.5; 0-1584, matching ST_LineWidth's 0-20,116,800
// EMU range). An out-of-range width is recorded as an error on the
// presentation (returned by Save) and leaves the border unset.
func (p *PictureRef) Border(c drawingml.Color, widthPoints float64) *PictureRef {
	p.spPr.Ln = newLn(p.pres, "Border", c, widthPoints)
	return p
}

// Shadow is ShapeRef.Shadow's counterpart for a placed image.
func (p *PictureRef) Shadow(c drawingml.Color, alphaPercent float64) *PictureRef {
	shdw, ok := newOuterShdw(p.pres, c, alphaPercent)
	if !ok {
		return p
	}
	effectLstOf(p.spPr).OuterShdw = shdw
	return p
}

// Glow is ShapeRef.Glow's counterpart for a placed image.
func (p *PictureRef) Glow(c drawingml.Color, radiusPoints float64) *PictureRef {
	glow, ok := newGlow(p.pres, c, radiusPoints)
	if !ok {
		return p
	}
	effectLstOf(p.spPr).Glow = glow
	return p
}

// Reflection is ShapeRef.Reflection's counterpart for a placed image.
func (p *PictureRef) Reflection(startOpacityPercent float64) *PictureRef {
	refl, ok := newReflection(p.pres, startOpacityPercent)
	if !ok {
		return p
	}
	effectLstOf(p.spPr).Reflection = refl
	return p
}

// SoftEdges is ShapeRef.SoftEdges's counterpart for a placed image.
func (p *PictureRef) SoftEdges(radiusPoints float64) *PictureRef {
	rad, ok := newSoftEdgeRad(p.pres, radiusPoints)
	if !ok {
		return p
	}
	effectLstOf(p.spPr).SoftEdge = &drawingml.SoftEdge{Rad: rad}
	return p
}
