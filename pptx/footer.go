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
	"strconv"

	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// slideNumberFieldGUID identifies the slide-number field instance (a:fld/@id,
// an ST_Guid). A fixed value is fine: the displayed number comes from the
// field's type, not its id, and PowerPoint itself reuses one guid for
// slide-number fields across a deck.
const slideNumberFieldGUID = "{F2E7E9A1-6C1F-4A5B-9C2D-1A2B3C4D5E6F}"

// Footer-row geometry, as percentages of the slide canvas: a short band along
// the bottom edge, split into date (left), footer (center), and slide number
// (right) — mirroring PowerPoint's own footer layout. Derived from the actual
// slide size so it tracks WithSlideSize / WithStandard4x3.
const (
	footerRowHeightPct  = 6  // band height
	footerRowBottomPct  = 3  // gap between the band and the bottom edge
	footerSideMarginPct = 4  // left/right page margin
	footerBandWidthPct  = 30 // date and footer band widths
	slideNumWidthPct    = 8  // slide-number band width
	footerFontPts       = 12 // footer-row text size, in points
)

// bottomPlaceholder builds and registers a footer-row placeholder of phType,
// self-positioned along the bottom edge at the given horizontal band (xPct,
// wPct as percentages of the slide width) with its own geometry — so it
// renders wherever the slide is shown, without depending on the slide master
// declaring a header/footer placeholder to inherit from. Returns the shape and
// true, or nil and false when the slide already has this footer-row
// placeholder (an error recorded on the presentation), mirroring
// AddPlaceholder's per-slide type+idx uniqueness guard.
func (s *Slide) bottomPlaceholder(phType PlaceholderType, idx uint32, xPct, wPct int) (*Shape, bool) {
	key := placeholderKey{phType: phType, idx: idx}
	if s.placeholders[key] {
		s.pres.addErr(errors.InvalidArgument("Footer/DateText/SlideNumber", "type", string(phType),
			"this slide already has this footer-row placeholder"))
		return nil, false
	}
	if s.placeholders == nil {
		s.placeholders = make(map[placeholderKey]bool)
	}
	s.placeholders[key] = true

	cx := s.pres.pres.SldSz.Cx
	cy := s.pres.pres.SldSz.Cy
	h := cy * footerRowHeightPct / 100
	y := cy - h - cy*footerRowBottomPct/100
	x := cx * xPct / 100
	w := cx * wPct / 100

	id := s.nextShapeID
	s.nextShapeID++
	shape := newPlaceholderShape(id, fmt.Sprintf("Placeholder %d", id), phType, idx, &drawingml.Xfrm{
		Off: &drawingml.Off{X: x, Y: y},
		Ext: &drawingml.Ext{Cx: w, Cy: h},
	})
	shape.SpPr.PrstGeom = &drawingml.PrstGeom{Prst: string(ShapeRect), AvLst: &drawingml.AvLst{}}
	s.spTree.Content = append(s.spTree.Content, shape)
	return shape, true
}

// footerParagraph appends a single formatted line of footer-row text to
// shape's text body at the given alignment.
func (s *Slide) footerParagraph(shape *Shape, text string, align Alignment) {
	para := &drawingml.Paragraph{}
	shape.TxBody.Paragraphs = append(shape.TxBody.Paragraphs, para)
	(&Paragraph{pres: s.pres, slidePath: s.path, p: para}).
		Text(text).FontSize(footerFontPts).ColorScheme(SchemeText1).Alignment(align)
}

// Footer places footer text along the bottom-center of the slide. It is a
// self-positioned ftr placeholder (with its own geometry, not inherited from
// the master), so it renders wherever the slide is shown. Calling Footer twice
// on one slide records an error (a duplicate placeholder), like AddPlaceholder.
func (s *Slide) Footer(text string) *Slide {
	if shape, ok := s.bottomPlaceholder(PlaceholderFooter, 11, (100-footerBandWidthPct)/2, footerBandWidthPct); ok {
		s.footerParagraph(shape, text, AlignCenter)
	}
	return s
}

// DateText places a date (or any short label) along the bottom-left of the
// slide, as a self-positioned dt placeholder. The text is literal — pptxgo
// writes exactly what you pass (deterministic, locale-independent) rather than
// an auto-updating date field. See Footer for the duplicate-call contract.
func (s *Slide) DateText(text string) *Slide {
	if shape, ok := s.bottomPlaceholder(PlaceholderDate, 10, footerSideMarginPct, footerBandWidthPct); ok {
		s.footerParagraph(shape, text, AlignLeft)
	}
	return s
}

// SlideNumber places an auto-updating slide-number field along the
// bottom-right of the slide: a sldNum placeholder holding an a:fld the
// consumer resolves to the slide's current position (so it stays correct if
// slides are reordered). The literal fallback is the slide's position at
// generation time. See Footer for the duplicate-call contract.
func (s *Slide) SlideNumber() *Slide {
	shape, ok := s.bottomPlaceholder(PlaceholderSlideNumber, 12, 100-footerSideMarginPct-slideNumWidthPct, slideNumWidthPct)
	if !ok {
		return s
	}
	shape.TxBody.Paragraphs = append(shape.TxBody.Paragraphs, &drawingml.Paragraph{
		PPr: &drawingml.PPr{Algn: string(AlignRight)},
		Content: []any{&drawingml.Fld{
			ID:   slideNumberFieldGUID,
			Type: "slidenum",
			RPr:  &drawingml.RPr{Sz: footerFontPts * 100},
			Text: strconv.Itoa(s.num),
		}},
	})
	return s
}
