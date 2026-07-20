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

// Inches converts f inches to EMUs. Shape positions and sizes (AddTextBox
// and friends) are all in EMUs; this and Points exist so call sites can
// name their unit instead of hand-computing the conversion.
func Inches(f float64) int {
	return int(f * drawingml.EMUsPerInch)
}

// Points converts f points to EMUs.
func Points(f float64) int {
	return int(f * drawingml.EMUsPerPoint)
}

// Emu returns n unchanged. It exists purely so a call site can spell out
// "this value is already in EMUs" instead of passing a bare int.
func Emu(n int) int {
	return n
}

// RGB constructs a drawingml.Color from 8-bit components, for use with
// Paragraph.Color.
func RGB(r, g, b uint8) drawingml.Color {
	return drawingml.Color{R: r, G: g, B: b}
}

// Alignment is a paragraph's horizontal text alignment (a:pPr's algn attribute).
type Alignment string

// Alignment values supported by Paragraph.Alignment.
const (
	AlignLeft    Alignment = "l"
	AlignCenter  Alignment = "ctr"
	AlignRight   Alignment = "r"
	AlignJustify Alignment = "just"
)

// SchemeColor references a color slot in the active theme's color scheme
// (a:schemeClr's val attribute, ST_SchemeColorVal — the slots the theme's
// own a:clrScheme defines) rather than an explicit RGB value, so a
// fill/border/text color automatically follows the theme. For use with
// ShapeRef.FillScheme, ShapeRef.BorderScheme, Paragraph.ColorScheme, and
// Slide.BackgroundScheme.
type SchemeColor string

// Theme color scheme slots. dk1/lt1/dk2/lt2 are the slots the theme's own
// a:clrScheme defines directly; bg1/tx1/bg2/tx2 are the same four slots
// under the aliases a slide's own p:clrMap maps them through (bg1->lt1,
// tx1->dk1, bg2->lt2, tx2->dk2, in pptxgo's default color map — see
// NewDefaultClrMap) — both forms are valid ST_SchemeColorVal values.
const (
	SchemeDark1             SchemeColor = "dk1"
	SchemeLight1            SchemeColor = "lt1"
	SchemeDark2             SchemeColor = "dk2"
	SchemeLight2            SchemeColor = "lt2"
	SchemeBackground1       SchemeColor = "bg1"
	SchemeText1             SchemeColor = "tx1"
	SchemeBackground2       SchemeColor = "bg2"
	SchemeText2             SchemeColor = "tx2"
	SchemeAccent1           SchemeColor = "accent1"
	SchemeAccent2           SchemeColor = "accent2"
	SchemeAccent3           SchemeColor = "accent3"
	SchemeAccent4           SchemeColor = "accent4"
	SchemeAccent5           SchemeColor = "accent5"
	SchemeAccent6           SchemeColor = "accent6"
	SchemeHyperlink         SchemeColor = "hlink"
	SchemeFollowedHyperlink SchemeColor = "folHlink"
)

// NumberingScheme names an automatic bullet-numbering scheme (a:buAutoNum's
// type attribute, ST_TextAutonumberScheme) for use with Paragraph.NumberedBullet.
type NumberingScheme string

// Common numbering schemes.
const (
	NumArabicPeriod  NumberingScheme = "arabicPeriod"  // "1.", "2.", ...
	NumArabicParenR  NumberingScheme = "arabicParenR"  // "1)", "2)", ...
	NumAlphaLcPeriod NumberingScheme = "alphaLcPeriod" // "a.", "b.", ...
	NumAlphaUcPeriod NumberingScheme = "alphaUcPeriod" // "A.", "B.", ...
	NumRomanLcPeriod NumberingScheme = "romanLcPeriod" // "i.", "ii.", ...
	NumRomanUcPeriod NumberingScheme = "romanUcPeriod" // "I.", "II.", ...
)

// VerticalAnchor is a text body's vertical anchoring within its shape
// (a:bodyPr's anchor attribute), for use with ShapeRef.Anchor.
type VerticalAnchor string

// Vertical anchor positions.
const (
	AnchorTop    VerticalAnchor = "t"
	AnchorMiddle VerticalAnchor = "ctr"
	AnchorBottom VerticalAnchor = "b"
)

// AutofitMode controls how a shape's text behaves when it overflows the
// shape's bounds, for use with ShapeRef.Autofit.
type AutofitMode string

// Autofit modes.
const (
	AutofitNone        AutofitMode = "none"  // text may overflow the shape uncorrected
	AutofitShrinkText  AutofitMode = "text"  // shrink font/line-spacing to fit
	AutofitResizeShape AutofitMode = "shape" // grow the shape to fit the text
)

// PlaceholderType names a placeholder's role (p:ph's type attribute,
// ST_PlaceholderType) — which same-typed, same-idx placeholder in a
// layout, and from there its master, a placeholder that omits its own
// position/formatting inherits from. Not exhaustive of ST_PlaceholderType's
// full set (which also names notes/date/footer/slide-number placeholders,
// among others) — these are the ones pptxgo's own master and standard
// layouts use.
type PlaceholderType string

// Placeholder types.
const (
	PlaceholderTitle       PlaceholderType = "title"    // main slide title
	PlaceholderCtrTitle    PlaceholderType = "ctrTitle" // centered title (title-slide layout)
	PlaceholderSubTitle    PlaceholderType = "subTitle" // subtitle (title-slide layout)
	PlaceholderBody        PlaceholderType = "body"     // bulleted body text
	PlaceholderDate        PlaceholderType = "dt"       // date, in the footer row (see Slide.DateText)
	PlaceholderFooter      PlaceholderType = "ftr"      // footer text (see Slide.Footer)
	PlaceholderSlideNumber PlaceholderType = "sldNum"   // slide-number field (see Slide.SlideNumber)
)

// GradientStop is one color stop within a linear gradient (see
// ShapeRef.GradientFill / Slide.BackgroundGradient): a color at a position
// along the gradient's axis. Pos is a percentage from 0 (the gradient's
// start) to 100 (its end) — supply stops in ascending Pos order for a
// well-formed gradient; nothing enforces that order itself.
//
// The stop's color is either an explicit RGB value (Color) or a theme color
// slot (Scheme). When Scheme is non-empty it takes precedence and the
// gradient stop follows the active theme (so a themed gradient recolors with
// WithTheme, just like FillScheme); leave Scheme empty ("") to use Color.
//
// Tint (0-100) lightens the stop's color toward white; Shade (0-100)
// darkens it toward black — at most one should be set (both are rarely
// meaningful together); zero on both means no adjustment, applied whether
// the stop uses Color or Scheme.
//
// Construct a GradientStop with keyed fields (e.g.
// GradientStop{Color: RGB(...), Pos: 0} or GradientStop{Scheme: SchemeAccent1,
// Pos: 0}), the form every call site here uses and that Go's vet composite
// check expects — Scheme, Tint, and Shade were all added as trailing
// optional fields, so keyed literals are unaffected.
type GradientStop struct {
	Color  drawingml.Color
	Pos    float64
	Scheme SchemeColor
	Tint   float64
	Shade  float64
}

// DashStyle names a preset line-dash pattern (a:prstDash's val attribute,
// ST_PresetLineDashVal) for use with ShapeRef.BorderDash.
type DashStyle string

// Preset dash patterns, the complete ST_PresetLineDashVal enumeration.
const (
	DashSolid         DashStyle = "solid"
	DashDot           DashStyle = "dot"
	DashDash          DashStyle = "dash"
	DashLgDash        DashStyle = "lgDash"
	DashDashDot       DashStyle = "dashDot"
	DashLgDashDot     DashStyle = "lgDashDot"
	DashLgDashDotDot  DashStyle = "lgDashDotDot"
	DashSysDash       DashStyle = "sysDash"
	DashSysDot        DashStyle = "sysDot"
	DashSysDashDot    DashStyle = "sysDashDot"
	DashSysDashDotDot DashStyle = "sysDashDotDot"
)

// validDashStyles is the complete ST_PresetLineDashVal enumeration.
var validDashStyles = map[DashStyle]bool{
	DashSolid: true, DashDot: true, DashDash: true, DashLgDash: true,
	DashDashDot: true, DashLgDashDot: true, DashLgDashDotDot: true,
	DashSysDash: true, DashSysDot: true, DashSysDashDot: true,
	DashSysDashDotDot: true,
}

// IsValidDashStyle reports whether style is one of ST_PresetLineDashVal's
// 11 defined preset dash pattern names.
func IsValidDashStyle(style DashStyle) bool {
	return validDashStyles[style]
}

// LineCapStyle names a line's end-cap style (a:ln's cap attribute,
// ST_LineCap), for use with ShapeRef.LineCap.
type LineCapStyle string

// The complete ST_LineCap enumeration.
const (
	LineCapFlat   LineCapStyle = "flat"
	LineCapRound  LineCapStyle = "rnd"
	LineCapSquare LineCapStyle = "sq"
)

// IsValidLineCapStyle reports whether style is one of ST_LineCap's 3
// defined values.
func IsValidLineCapStyle(style LineCapStyle) bool {
	return style == LineCapFlat || style == LineCapRound || style == LineCapSquare
}

// LineJoinStyle names a line's corner-join style — how two of its segments
// meet at a corner — for use with ShapeRef.LineJoin.
type LineJoinStyle string

// The three EG_LineJoinProperties choices.
const (
	LineJoinRound LineJoinStyle = "round"
	LineJoinBevel LineJoinStyle = "bevel"
	LineJoinMiter LineJoinStyle = "miter" // Office's own default miter limit, 800%
)

// ArrowheadType names an arrowhead (or other line-end decoration) for use
// with ShapeRef.ArrowStart/ArrowEnd (a:headEnd/a:tailEnd's type attribute,
// ST_LineEndType). Only has visible effect on an open shape's outline (e.g.
// ShapeLine) — a closed autoshape's path has no defined start/end.
type ArrowheadType string

// The complete ST_LineEndType enumeration.
const (
	ArrowheadNone     ArrowheadType = "none"
	ArrowheadTriangle ArrowheadType = "triangle"
	ArrowheadStealth  ArrowheadType = "stealth"
	ArrowheadDiamond  ArrowheadType = "diamond"
	ArrowheadOval     ArrowheadType = "oval"
	ArrowheadArrow    ArrowheadType = "arrow"
)

// validArrowheadTypes is the complete ST_LineEndType enumeration.
var validArrowheadTypes = map[ArrowheadType]bool{
	ArrowheadNone: true, ArrowheadTriangle: true, ArrowheadStealth: true,
	ArrowheadDiamond: true, ArrowheadOval: true, ArrowheadArrow: true,
}

// IsValidArrowheadType reports whether t is one of ST_LineEndType's 6
// defined values.
func IsValidArrowheadType(t ArrowheadType) bool {
	return validArrowheadTypes[t]
}

// TableCellSide names which edge (or diagonal) of a table cell
// TableCell.Border/BorderScheme sets an outline on (a:tcPr's six per-side
// line children).
type TableCellSide string

// The six CT_TableCellProperties line-child sides.
const (
	SideLeft         TableCellSide = "l"
	SideRight        TableCellSide = "r"
	SideTop          TableCellSide = "t"
	SideBottom       TableCellSide = "b"
	SideDiagonalDown TableCellSide = "tlToBr" // top-left to bottom-right
	SideDiagonalUp   TableCellSide = "blToTr" // bottom-left to top-right
)

// PresetGeometry names a preset autoshape outline (a:prstGeom's prst
// attribute, schema type ST_ShapeType) for use with Slide.AddShape. This is
// a representative subset of the ~180 shapes ST_ShapeType allows; any other
// valid preset name can still be passed as a plain PresetGeometry("name").
type PresetGeometry string

// Common preset geometries.
const (
	ShapeLine           PresetGeometry = "line"
	ShapeRect           PresetGeometry = "rect"
	ShapeRoundRect      PresetGeometry = "roundRect"
	ShapeEllipse        PresetGeometry = "ellipse"
	ShapeTriangle       PresetGeometry = "triangle"
	ShapeRightTriangle  PresetGeometry = "rtTriangle"
	ShapeParallelogram  PresetGeometry = "parallelogram"
	ShapeTrapezoid      PresetGeometry = "trapezoid"
	ShapeDiamond        PresetGeometry = "diamond"
	ShapePentagon       PresetGeometry = "pentagon"
	ShapeHexagon        PresetGeometry = "hexagon"
	ShapeHeptagon       PresetGeometry = "heptagon"
	ShapeOctagon        PresetGeometry = "octagon"
	ShapeStar4          PresetGeometry = "star4"
	ShapeStar5          PresetGeometry = "star5"
	ShapeStar6          PresetGeometry = "star6"
	ShapeStar8          PresetGeometry = "star8"
	ShapeRightArrow     PresetGeometry = "rightArrow"
	ShapeLeftArrow      PresetGeometry = "leftArrow"
	ShapeUpArrow        PresetGeometry = "upArrow"
	ShapeDownArrow      PresetGeometry = "downArrow"
	ShapeLeftRightArrow PresetGeometry = "leftRightArrow"
	ShapeUpDownArrow    PresetGeometry = "upDownArrow"
	ShapeChevron        PresetGeometry = "chevron"
	ShapeDonut          PresetGeometry = "donut"
	ShapeNoSmoking      PresetGeometry = "noSmoking"
	ShapeHeart          PresetGeometry = "heart"
	ShapeLightningBolt  PresetGeometry = "lightningBolt"
	ShapeSun            PresetGeometry = "sun"
	ShapeMoon           PresetGeometry = "moon"
	ShapeCloud          PresetGeometry = "cloud"
	ShapeArc            PresetGeometry = "arc"
	ShapePlaque         PresetGeometry = "plaque"
	ShapeCan            PresetGeometry = "can"
	ShapeCube           PresetGeometry = "cube"
	ShapeBevel          PresetGeometry = "bevel"
	ShapeSmileyFace     PresetGeometry = "smileyFace"
	ShapeWave           PresetGeometry = "wave"
	ShapeDoubleWave     PresetGeometry = "doubleWave"
)

// ConnectorType names a connector's own line geometry (a:prstGeom's prst
// attribute on a p:cxnSp, the same ST_ShapeType namespace AddShape's
// PresetGeometry draws from — but only the connector-shaped subset makes
// sense on a p:cxnSp), for use with Slide.Connect.
type ConnectorType string

// Common connector geometries. Any other connector-shaped ST_ShapeType name
// (e.g. "curvedConnector2") can still be passed as a plain
// ConnectorType("name") — this is a representative subset, the same
// convention PresetGeometry's own named constants use.
const (
	ConnStraight ConnectorType = "line"             // a direct line between the two connection sites
	ConnBent     ConnectorType = "bentConnector3"   // right-angled routing, PowerPoint's own default connector style
	ConnCurved   ConnectorType = "curvedConnector3" // curved routing
)

// ConnSite names a connection site on a shape's own geometry, for use with
// Slide.Connect. Built-in autoshapes (rect, roundRect, ellipse, and the
// other common diagram shapes) number their four cardinal connection sites
// 0 (top), 1 (left), 2 (bottom), 3 (right), counter-clockwise from the top —
// see drawingml.StCxn's own doc comment for how this was confirmed against
// a real render rather than assumed from the schema.
type ConnSite string

// The four cardinal connection sites shared by rect/roundRect/ellipse and
// most other built-in autoshapes.
const (
	SiteTop    ConnSite = "top"
	SiteLeft   ConnSite = "left"
	SiteBottom ConnSite = "bottom"
	SiteRight  ConnSite = "right"
)

// connSiteIdx maps the four cardinal ConnSite names to their
// ST_ShapeType-standard connection-site index (see ConnSite's own doc
// comment).
var connSiteIdx = map[ConnSite]uint32{
	SiteTop: 0, SiteLeft: 1, SiteBottom: 2, SiteRight: 3,
}
