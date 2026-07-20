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
	"fmt"

	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// Connector is p:cxnSp (CT_Connector): a line shape whose ends can be bound
// to connection sites on other shapes (see Slide.Connect) — unlike an
// ordinary autoshape's a:ln outline (ShapeRef.Border and friends), a bound
// connector's endpoints move with the shapes they're attached to when
// PowerPoint's own UI repositions them. Reuses *SpPr (the same xfrm/
// prstGeom/ln every p:sp already carries) but has no txBody — CT_Connector
// itself defines none, so unlike Shape, a connector can never hold text.
// Field order mirrors the schema: nvCxnSpPr -> spPr.
type Connector struct {
	XMLName   xml.Name   `xml:"p:cxnSp"`
	NvCxnSpPr *NvCxnSpPr `xml:"p:nvCxnSpPr"`
	SpPr      *SpPr      `xml:"p:spPr"`
}

// NvCxnSpPr is p:nvCxnSpPr (CT_ConnectorNonVisual): a connector's non-visual
// properties — the same cNvPr/nvPr shape every other shape-like element
// carries (see NvSpPr, NvPicPr), plus cNvCxnSpPr for the connection bindings.
type NvCxnSpPr struct {
	XMLName    xml.Name    `xml:"p:nvCxnSpPr"`
	CNvPr      *CNvPr      `xml:"p:cNvPr"`
	CNvCxnSpPr *CNvCxnSpPr `xml:"p:cNvCxnSpPr"`
	NvPr       *NvPr       `xml:"p:nvPr"`
}

// CNvCxnSpPr is p:cNvCxnSpPr (CT_NonVisualConnectorProperties):
// connector-specific non-visual drawing properties. StCxn/EndCxn bind the
// connector's start/end points to another shape's connection site — see
// drawingml.StCxn's own doc comment for the index convention. CxnSpLocks
// (restricting move/resize in the authoring UI) is out of scope until a
// caller needs it.
type CNvCxnSpPr struct {
	XMLName xml.Name          `xml:"p:cNvCxnSpPr"`
	StCxn   *drawingml.StCxn  `xml:"a:stCxn,omitempty"`
	EndCxn  *drawingml.EndCxn `xml:"a:endCxn,omitempty"`
}

// ConnectorRef is a handle onto a placed connector (a p:cxnSp), returned by
// Slide.Connect. It exposes the same line-styling methods ShapeRef does
// (Border, BorderScheme, BorderDash, LineCap, LineJoin, ArrowStart,
// ArrowEnd) via the shared apply*/build* helpers in text_builder.go — but
// is its own type, not a ShapeRef alias, since a connector has no txBody
// (see Connector's own doc comment) for AddParagraph and the other
// text-formatting methods to target.
type ConnectorRef struct {
	pres *Presentation
	spPr *SpPr
}

// Border sets the connector's line to a solid color at the given width, in
// points — see ShapeRef.Border for the width's valid range and error
// behavior.
func (cr *ConnectorRef) Border(c drawingml.Color, widthPoints float64) *ConnectorRef {
	cr.spPr.Ln = newLn(cr.pres, "Border", c, widthPoints)
	return cr
}

// BorderScheme is Border's theme-color counterpart, referencing a scheme
// slot (e.g. SchemeAccent1) rather than an explicit RGB value.
func (cr *ConnectorRef) BorderScheme(scheme SchemeColor, widthPoints float64) *ConnectorRef {
	cr.spPr.Ln = newLnScheme(cr.pres, "BorderScheme", scheme, widthPoints)
	return cr
}

// BorderDash sets the connector's line to a preset dash pattern — see
// ShapeRef.BorderDash for the prior-Border requirement and error behavior.
func (cr *ConnectorRef) BorderDash(style DashStyle) *ConnectorRef {
	applyBorderDash(cr.pres, cr.spPr, "BorderDash", style)
	return cr
}

// LineCap sets the connector's own end-cap style — see ShapeRef.LineCap for
// the prior-Border requirement and error behavior.
func (cr *ConnectorRef) LineCap(style LineCapStyle) *ConnectorRef {
	applyLineCap(cr.pres, cr.spPr, "LineCap", style)
	return cr
}

// LineJoin sets the connector's own corner-join style (visible on a
// ConnBent/ConnCurved connector's own routed corners, not ConnStraight) —
// see ShapeRef.LineJoin for the prior-Border requirement and error behavior.
func (cr *ConnectorRef) LineJoin(style LineJoinStyle) *ConnectorRef {
	applyLineJoin(cr.pres, cr.spPr, "LineJoin", style)
	return cr
}

// ArrowStart sets an arrowhead at the connector's own bound start point —
// see ShapeRef.ArrowStart for the prior-Border requirement and error
// behavior.
func (cr *ConnectorRef) ArrowStart(t ArrowheadType) *ConnectorRef {
	end, ok := buildArrowEnd(cr.pres, cr.spPr, "ArrowStart", t)
	if !ok {
		return cr
	}
	cr.spPr.Ln.HeadEnd = end
	return cr
}

// ArrowEnd is ArrowStart's counterpart for the connector's own bound end
// point.
func (cr *ConnectorRef) ArrowEnd(t ArrowheadType) *ConnectorRef {
	end, ok := buildArrowEnd(cr.pres, cr.spPr, "ArrowEnd", t)
	if !ok {
		return cr
	}
	cr.spPr.Ln.TailEnd = end
	return cr
}

// Connect adds a connector (p:cxnSp) whose start and end points are BOUND
// to connection sites on from and to (see ConnSite) — in PowerPoint's own
// UI, moving either shape re-routes the connector to follow, the behavior
// that distinguishes a real connector from a plain line shape
// (Slide.AddShape(ShapeLine, ...), which has no such binding). ct selects
// the connector's own routing geometry (e.g. ConnStraight, ConnBent — the
// default when ct is unset would be the zero value "", which is not a
// valid ST_ShapeType name, so callers must pass one explicitly).
//
// The emitted a:xfrm spans the two connection POINTS the connector binds
// (not a bounding box of the two shapes' full rectangles — see connectorXfrm
// for why that distinction is load-bearing), computed from each endpoint's
// own spPr.Xfrm. A placeholder has no a:xfrm of its own (it inherits
// geometry from the layout — see AddPlaceholder), so Connect records an error
// on the presentation and leaves the connector unset when either endpoint is
// one; a shape added inside a Group is NOT such a case (Group.AddShape gives
// it a slide-absolute a:xfrm, so group members are connectable — the demo
// binds across a group boundary). PowerPoint recomputes the connector's
// actual visual routing from the binding once opened; the a:xfrm only needs
// to be non-degenerate, not an exact fit.
//
// Both endpoints must be drawn with a preset in connSiteGeom (rect,
// roundRect, ellipse) — the geometries whose connection-site indices are
// verified (see ConnSite). An endpoint with any other preset, an
// unrecognized fromSite/toSite, or an invalid ct is recorded as an error on
// the presentation (returned by Save) and leaves the connector unset.
//
// siteXY reads each endpoint's un-rotated cardinal point, so a connector to a
// shape that has been rotated (ShapeRef.Rotation) or flipped (FlipH/FlipV)
// may render slightly detached on first paint; because the stCxn/endCxn
// binding is still correct, PowerPoint re-routes it to the shape the moment
// that shape is nudged. Rotation-aware endpoints are a possible follow-up.
func (s *Slide) Connect(from *ShapeRef, fromSite ConnSite, to *ShapeRef, toSite ConnSite, ct ConnectorType) *ConnectorRef {
	fromIdx, ok := connSiteIdx[fromSite]
	if !ok {
		s.pres.addErr(errors.InvalidArgument("Connect", "fromSite", fromSite, "not a valid connection site"))
		return &ConnectorRef{pres: s.pres, spPr: &SpPr{}}
	}
	toIdx, ok := connSiteIdx[toSite]
	if !ok {
		s.pres.addErr(errors.InvalidArgument("Connect", "toSite", toSite, "not a valid connection site"))
		return &ConnectorRef{pres: s.pres, spPr: &SpPr{}}
	}
	if !IsValidPresetGeometry(PresetGeometry(ct)) {
		s.pres.addErr(errors.InvalidArgument("Connect", "ct", ct, "must be a valid ST_ShapeType connector geometry (e.g. ConnStraight, ConnBent)"))
		return &ConnectorRef{pres: s.pres, spPr: &SpPr{}}
	}
	if from.spPr.Xfrm == nil || to.spPr.Xfrm == nil {
		s.pres.addErr(errors.InvalidArgument("Connect", "from/to", "no xfrm",
			"both shapes need their own a:xfrm to connect (a placeholder inherits geometry from the layout and has none — see AddPlaceholder)"))
		return &ConnectorRef{pres: s.pres, spPr: &SpPr{}}
	}
	if !connSiteBindable(from) || !connSiteBindable(to) {
		s.pres.addErr(errors.InvalidArgument("Connect", "from/to", "unsupported geometry",
			"both shapes must be drawn with rect, roundRect, or ellipse — the presets whose connection-site indices are verified (see ConnSite)"))
		return &ConnectorRef{pres: s.pres, spPr: &SpPr{}}
	}

	id := s.allocID()
	fromX, fromY := siteXY(from.spPr.Xfrm, fromSite)
	toX, toY := siteXY(to.spPr.Xfrm, toSite)
	spPr := &SpPr{
		Xfrm:     connectorXfrm(fromX, fromY, toX, toY),
		PrstGeom: &drawingml.PrstGeom{Prst: string(ct), AvLst: &drawingml.AvLst{}},
	}
	conn := &Connector{
		NvCxnSpPr: &NvCxnSpPr{
			CNvPr: &CNvPr{ID: id, Name: fmt.Sprintf("Connector %d", id)},
			CNvCxnSpPr: &CNvCxnSpPr{
				StCxn:  &drawingml.StCxn{ID: from.id, Idx: fromIdx},
				EndCxn: &drawingml.EndCxn{ID: to.id, Idx: toIdx},
			},
			NvPr: &NvPr{},
		},
		SpPr: spPr,
	}
	s.spTree.Content = append(s.spTree.Content, conn)

	return &ConnectorRef{pres: s.pres, spPr: spPr}
}

// connSiteBindable reports whether sr's shape is drawn with a preset geometry
// whose connection-site indices connSiteIdx is verified against (see
// connSiteGeom). A shape with no preset geometry at all is not bindable.
func connSiteBindable(sr *ShapeRef) bool {
	if sr.spPr.PrstGeom == nil {
		return false
	}
	return connSiteGeom[PresetGeometry(sr.spPr.PrstGeom.Prst)]
}

// siteXY returns the slide-absolute (x, y) of the given connection site on
// a shape whose own transform is xfrm — the exact point PowerPoint's own
// "line" connector geometry draws to/from, not merely a point somewhere on
// the shape. It reads the site off the un-rotated, un-flipped box (Off/Ext
// only, ignoring xfrm.Rot/FlipH/FlipV); for a rotated or flipped endpoint the
// first-paint point is therefore approximate, self-correcting once PowerPoint
// re-routes from the binding (see Connect's own doc comment). The four
// cardinal offsets match ConnSite's own doc comment
// (0=top, 1=left, 2=bottom, 3=right, counter-clockwise from the top),
// confirmed against python-pptx's own hardcoded connection-point formula
// (Connector._move_begin_to_cxn/_move_end_to_cxn) and a real render — the
// same "extract from a real, working implementation" discipline
// drawingml.StCxn's own doc comment already documents.
func siteXY(xfrm *drawingml.Xfrm, site ConnSite) (x, y int) {
	ox, oy := xfrm.Off.X, xfrm.Off.Y
	cx, cy := xfrm.Ext.Cx, xfrm.Ext.Cy
	switch site {
	case SiteTop:
		return ox + cx/2, oy
	case SiteLeft:
		return ox, oy + cy/2
	case SiteBottom:
		return ox + cx/2, oy + cy
	case SiteRight:
		return ox + cx, oy + cy/2
	}
	return ox, oy // unreachable: Connect validates site via connSiteIdx first
}

// connectorXfrm builds a connector's own a:xfrm spanning the two connection
// POINTS it binds (fromX, fromY) -> (toX, toY) — NOT a bounding box of the
// two shapes' full rectangles. This distinction is the fix for a real bug:
// an earlier version bounding-boxed the two shapes themselves, so two
// same-height boxes at the same y produced a box with a non-zero height —
// and a:prstGeom "line" draws the LITERAL diagonal of whatever box it's
// given (it has no routing logic of its own, unlike bentConnector/
// curvedConnector's own preset formulas), so PowerPoint's first paint drew
// a diagonal cutting through both shapes instead of the intended flat
// line — confirmed only by opening in real PowerPoint (a static LibreOffice
// render happened to look fine, and the diagonal in PowerPoint itself
// self-corrected the moment either connected shape was dragged, since that
// forces PowerPoint to recompute the connector's actual routing from its
// stCxn/endCxn binding) — the same "SDK-valid, schema-legal, still wrong"
// class of defect as OuterShdw's own missing sy fix.
//
// ST_PositiveSize2D forbids a negative Ext, so a connector whose end point
// is above/left of its start point needs FlipV/FlipH instead — matching
// python-pptx's own begin_x/end_x accessors: begin is the flipped corner,
// end is the un-flipped one, so a "from right of to" connection correctly
// draws right-to-left rather than silently reversing which point is bound
// to which shape.
func connectorXfrm(fromX, fromY, toX, toY int) *drawingml.Xfrm {
	flipH := fromX > toX
	flipV := fromY > toY
	minX, maxX := fromX, toX
	if minX > maxX {
		minX, maxX = maxX, minX
	}
	minY, maxY := fromY, toY
	if minY > maxY {
		minY, maxY = maxY, minY
	}
	return &drawingml.Xfrm{
		FlipH: flipH,
		FlipV: flipV,
		Off:   &drawingml.Off{X: minX, Y: minY},
		Ext:   &drawingml.Ext{Cx: maxX - minX, Cy: maxY - minY},
	}
}
