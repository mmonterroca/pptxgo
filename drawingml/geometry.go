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

// Package drawingml implements the DrawingML (ECMA-376 "a:" namespace)
// primitives shared by DOCX, PPTX, and XLSX: shape transforms, preset
// geometry, colors, fills, and image references.
//
// Everything in this package uses only the "a:" namespace. Deliberately
// excluded is any wrapper that fixes a namespace prefix to one host format —
// most notably the picture container ("pic:pic" when a DOCX paragraph
// embeds a graphic, "p:pic" when a PPTX slide places one directly). Both
// wrap the exact same shape — non-visual properties, a blip fill, shape
// properties — but under a different element prefix, so that container
// belongs to the consuming package, built out of the types defined here.
package drawingml

import "encoding/xml"

// Xfrm is a 2D transform (a:xfrm): position and size, in EMUs, with
// optional rotation and flip.
type Xfrm struct {
	XMLName xml.Name `xml:"a:xfrm"`
	Rot     int      `xml:"rot,attr,omitempty"` // rotation, 60,000ths of a degree
	FlipH   bool     `xml:"flipH,attr,omitempty"`
	FlipV   bool     `xml:"flipV,attr,omitempty"`
	Off     *Off     `xml:"a:off,omitempty"`
	Ext     *Ext     `xml:"a:ext"`
}

// Off is a position offset (a:off), in EMUs.
type Off struct {
	XMLName xml.Name `xml:"a:off"`
	X       int      `xml:"x,attr"`
	Y       int      `xml:"y,attr"`
}

// Ext is a size extent (a:ext), in EMUs.
type Ext struct {
	XMLName xml.Name `xml:"a:ext"`
	Cx      int      `xml:"cx,attr"`
	Cy      int      `xml:"cy,attr"`
}

// PrstGeom is a preset shape geometry (a:prstGeom), e.g. "rect", "ellipse".
type PrstGeom struct {
	XMLName xml.Name `xml:"a:prstGeom"`
	Prst    string   `xml:"prst,attr"`
	AvLst   *AvLst   `xml:"a:avLst,omitempty"`
}

// AvLst is an adjust-value list (a:avLst) for a preset geometry's
// parametric handles. Empty for shapes using their default proportions.
type AvLst struct {
	XMLName xml.Name `xml:"a:avLst"`
	Gd      []*Gd    `xml:"a:gd,omitempty"`
}

// Gd is a single named guide/adjust value (a:gd) within an AvLst.
type Gd struct {
	XMLName xml.Name `xml:"a:gd"`
	Name    string   `xml:"name,attr"`
	Fmla    string   `xml:"fmla,attr"`
}

// GraphicFrameLocks (a:graphicFrameLocks) restricts what a user may do to a
// graphic frame (a table, chart, or embedded object) in the authoring UI.
type GraphicFrameLocks struct {
	XMLName        xml.Name `xml:"a:graphicFrameLocks"`
	NoChangeAspect bool     `xml:"noChangeAspect,attr,omitempty"`
	NoMove         bool     `xml:"noMove,attr,omitempty"`
	NoResize       bool     `xml:"noResize,attr,omitempty"`
}

// GroupXfrm is a group shape's own a:xfrm (CT_GroupTransform2D) — distinct
// from Xfrm (CT_Transform2D, an ordinary shape's own transform): both
// happen to share the "a:xfrm" element name (used in different parent
// contexts, p:grpSpPr vs. p:spPr, so there is no field-tag conflict — see
// ChOff's own doc comment for the conflict this design DOES have to dodge),
// but only GroupXfrm carries ChOff/ChExt, the child coordinate space every
// shape nested inside the group is positioned in. Field order mirrors the
// schema: off, ext, chOff, chExt.
type GroupXfrm struct {
	XMLName xml.Name `xml:"a:xfrm"`
	Rot     int      `xml:"rot,attr,omitempty"` // rotation, 60,000ths of a degree
	FlipH   bool     `xml:"flipH,attr,omitempty"`
	FlipV   bool     `xml:"flipV,attr,omitempty"`
	Off     *Off     `xml:"a:off,omitempty"`
	Ext     *Ext     `xml:"a:ext,omitempty"`
	ChOff   *ChOff   `xml:"a:chOff,omitempty"`
	ChExt   *ChExt   `xml:"a:chExt,omitempty"`
}

// ChOff is a:chOff (CT_Point2D): the top-left corner of the coordinate
// space a group's CHILDREN are positioned in — paired with ChExt to
// complete GroupXfrm. A separate type from Off, not a reuse of it, even
// though both are schema-identical CT_Point2D: Off's own fixed "a:off"
// XMLName would win over any field tag that tried to rename it to
// "a:chOff" — the same conflict TcBorderLn/LineEnd/GraphicFrameXfrm's own
// doc comments already document, encoding/xml's rule that a nested type's
// own XMLName always wins.
//
// Slide.AddGroup sets ChOff equal to the group's own Off and ChExt equal to
// its own Ext (a 1:1 child space) so a shape added inside the group at a
// given (x, y) lands at that exact slide position — the same coordinates
// it would use outside the group. The general mapping PowerPoint applies is
// p_parent = off + (p_child − chOff)·(ext/chExt); chOff=off ∧ chExt=ext
// makes that the identity function.
type ChOff struct {
	XMLName xml.Name `xml:"a:chOff"`
	X       int      `xml:"x,attr"`
	Y       int      `xml:"y,attr"`
}

// ChExt is a:chExt (CT_PositiveSize2D): the size of a group's child
// coordinate space — see ChOff's own doc comment for the full mapping and
// why this can't reuse Ext directly.
type ChExt struct {
	XMLName xml.Name `xml:"a:chExt"`
	Cx      int      `xml:"cx,attr"`
	Cy      int      `xml:"cy,attr"`
}
