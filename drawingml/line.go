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

package drawingml

import "encoding/xml"

// Ln is a:ln (CT_LineProperties): a shape or picture's outline (or, used
// inside a table cell's a:tcPr, one border side). Field order mirrors the
// schema: cap is an attribute alongside w; the fill group (SolidFill)
// precedes PrstDash, which precedes the line-join choice (Round/Bevel/Miter —
// at most one should ever be set, the same EG_LineJoinProperties
// mutual-exclusion pattern the fill group elsewhere in this package
// follows), which precedes the arrowhead ends HeadEnd/TailEnd.
type Ln struct {
	XMLName   xml.Name   `xml:"a:ln"`
	W         int        `xml:"w,attr,omitempty"`   // EMUs
	Cap       string     `xml:"cap,attr,omitempty"` // ST_LineCap: flat/rnd/sq
	SolidFill *SolidFill `xml:"a:solidFill,omitempty"`
	PrstDash  *PrstDash  `xml:"a:prstDash,omitempty"`
	Round     *LnRound   `xml:"a:round,omitempty"`
	Bevel     *LnBevel   `xml:"a:bevel,omitempty"`
	Miter     *LnMiter   `xml:"a:miter,omitempty"`
	HeadEnd   *LineEnd   `xml:"a:headEnd,omitempty"`
	TailEnd   *LineEnd   `xml:"a:tailEnd,omitempty"`
}

// PrstDash is a:prstDash (CT_PresetLineDashProperties): a named dash pattern
// for a line's outline, referencing one of ST_PresetLineDashVal's values
// (e.g. "solid", "dash", "dot", "dashDot", "sysDash").
type PrstDash struct {
	XMLName xml.Name `xml:"a:prstDash"`
	Val     string   `xml:"val,attr"`
}

// LnRound is a:round (CT_LineJoinRound): a rounded corner where two line
// segments meet. Always empty.
type LnRound struct {
	XMLName xml.Name `xml:"a:round"`
}

// LnBevel is a:bevel (CT_LineJoinBevel): a flattened (cut-off) corner where
// two line segments meet. Always empty.
type LnBevel struct {
	XMLName xml.Name `xml:"a:bevel"`
}

// LnMiter is a:miter (CT_LineJoinMiterProperties): a sharp, pointed corner
// where two line segments meet, clipped once it extends past Lim (in
// thousandths of a percent of the line width) — Office's own built-in line
// styles use 800000 (800%, see themeFmtScheme).
type LnMiter struct {
	XMLName xml.Name `xml:"a:miter"`
	Lim     int      `xml:"lim,attr,omitempty"`
}

// LineEnd is a:headEnd/a:tailEnd (CT_LineEndProperties): an arrowhead or
// other decoration at one end of an open line's path (a closed autoshape's
// outline has no defined start/end, so this only has visible effect on an
// open shape, e.g. prstGeom "line"). Deliberately has no XMLName of its own
// — unlike every other element type in this package — so the same struct
// serializes as either a:headEnd or a:tailEnd purely from the containing
// field's own xml tag; giving it a fixed XMLName would make that tag always
// win instead (the same gotcha GraphicFrameXfrm's own doc comment records
// for a:xfrm). Type names the decoration (ST_LineEndType: none/triangle/
// stealth/diamond/oval/arrow); W and Len scale it (ST_LineEndWidth/
// ST_LineEndLength: sm/med/lg).
type LineEnd struct {
	Type string `xml:"type,attr,omitempty"`
	W    string `xml:"w,attr,omitempty"`
	Len  string `xml:"len,attr,omitempty"`
}

// NewLn returns a solid-color outline of the given width in EMUs.
func NewLn(c Color, widthEMU int) *Ln {
	return &Ln{W: widthEMU, SolidFill: NewSolidFillRGB(c)}
}

// NewLnScheme returns a theme-color outline of the given width in EMUs.
func NewLnScheme(schemeColor string, widthEMU int) *Ln {
	return &Ln{W: widthEMU, SolidFill: NewSolidFillScheme(schemeColor)}
}
