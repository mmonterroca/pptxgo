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

import (
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// Color is an RGB color.
type Color struct {
	R, G, B uint8
}

// Common color constants for convenience.
var (
	Black = Color{R: 0, G: 0, B: 0}
	White = Color{R: 255, G: 255, B: 255}
)

// ToHex renders c as a 6-digit hex string with no leading "#" (e.g. "FF0000"),
// the form DrawingML's a:srgbClr val attribute expects.
func ToHex(c Color) string {
	return fmt.Sprintf("%02X%02X%02X", c.R, c.G, c.B)
}

// FromHex parses a hex color string. Accepts "RGB", "RRGGBB", "#RGB", "#RRGGBB".
func FromHex(hex string) (Color, error) {
	if len(hex) > 0 && hex[0] == '#' {
		hex = hex[1:]
	}

	switch len(hex) {
	case 3:
		rv, err := strconv.ParseUint(string(hex[0]), 16, 8)
		if err != nil {
			return Color{}, errors.InvalidArgument("FromHex", "hex", hex, "invalid red component")
		}
		gv, err := strconv.ParseUint(string(hex[1]), 16, 8)
		if err != nil {
			return Color{}, errors.InvalidArgument("FromHex", "hex", hex, "invalid green component")
		}
		bv, err := strconv.ParseUint(string(hex[2]), 16, 8)
		if err != nil {
			return Color{}, errors.InvalidArgument("FromHex", "hex", hex, "invalid blue component")
		}
		return Color{R: uint8(rv*16 + rv), G: uint8(gv*16 + gv), B: uint8(bv*16 + bv)}, nil

	case 6:
		rv, err := strconv.ParseUint(hex[0:2], 16, 8)
		if err != nil {
			return Color{}, errors.InvalidArgument("FromHex", "hex", hex, "invalid red component")
		}
		gv, err := strconv.ParseUint(hex[2:4], 16, 8)
		if err != nil {
			return Color{}, errors.InvalidArgument("FromHex", "hex", hex, "invalid green component")
		}
		bv, err := strconv.ParseUint(hex[4:6], 16, 8)
		if err != nil {
			return Color{}, errors.InvalidArgument("FromHex", "hex", hex, "invalid blue component")
		}
		return Color{R: uint8(rv), G: uint8(gv), B: uint8(bv)}, nil

	default:
		return Color{}, errors.InvalidArgument("FromHex", "hex", hex,
			"hex color must be 3 or 6 characters (optionally prefixed with #)")
	}
}

// SolidFill (a:solidFill) fills with a single color, given as either an
// explicit RGB value (SrgbClr) or a reference into the active theme's color
// scheme (SchemeClr). Exactly one should be set.
type SolidFill struct {
	XMLName   xml.Name   `xml:"a:solidFill"`
	SrgbClr   *SrgbClr   `xml:"a:srgbClr,omitempty"`
	SchemeClr *SchemeClr `xml:"a:schemeClr,omitempty"`
}

// NewSolidFillRGB creates a SolidFill from an explicit color.
func NewSolidFillRGB(c Color) *SolidFill {
	return &SolidFill{SrgbClr: &SrgbClr{Val: ToHex(c)}}
}

// NewSolidFillScheme creates a SolidFill referencing a theme color slot
// (e.g. "accent1", "dk1", "lt1" — see the theme's a:clrScheme).
func NewSolidFillScheme(schemeColor string) *SolidFill {
	return &SolidFill{SchemeClr: &SchemeClr{Val: schemeColor}}
}

// SrgbClr (a:srgbClr) is an explicit RGB color, as a 6-digit hex string,
// optionally adjusted by one or more color transforms (Tint/Shade/Alpha/
// LumMod/LumOff — see their own doc comments). The schema's EG_ColorTransform
// group is a repeatable *choice* (each transform 0 or more times, in any
// order), so unlike most sibling elements in this package, field order here
// is stylistic, not schema-mandated — real Office output (see
// themeFmtScheme) shows tint before shade before lumMod, which this mirrors.
type SrgbClr struct {
	XMLName xml.Name `xml:"a:srgbClr"`
	Val     string   `xml:"val,attr"`
	Tint    *Tint    `xml:"a:tint,omitempty"`
	Shade   *Shade   `xml:"a:shade,omitempty"`
	Alpha   *Alpha   `xml:"a:alpha,omitempty"`
	LumMod  *LumMod  `xml:"a:lumMod,omitempty"`
	LumOff  *LumOff  `xml:"a:lumOff,omitempty"`
}

// SchemeClr (a:schemeClr) references a color slot from the active theme's
// color scheme, optionally adjusted by the same color transforms SrgbClr
// carries — see SrgbClr's doc comment for the field-order note.
type SchemeClr struct {
	XMLName xml.Name `xml:"a:schemeClr"`
	Val     string   `xml:"val,attr"`
	Tint    *Tint    `xml:"a:tint,omitempty"`
	Shade   *Shade   `xml:"a:shade,omitempty"`
	Alpha   *Alpha   `xml:"a:alpha,omitempty"`
	LumMod  *LumMod  `xml:"a:lumMod,omitempty"`
	LumOff  *LumOff  `xml:"a:lumOff,omitempty"`
}

// Tint is a:tint (CT_PositiveFixedPercentage): lightens a color toward
// white by the given percentage, in thousandths of a percent (0-100000).
type Tint struct {
	XMLName xml.Name `xml:"a:tint"`
	Val     int      `xml:"val,attr"`
}

// Shade is a:shade (CT_PositiveFixedPercentage): darkens a color toward
// black by the given percentage, in thousandths of a percent (0-100000).
type Shade struct {
	XMLName xml.Name `xml:"a:shade"`
	Val     int      `xml:"val,attr"`
}

// Alpha is a:alpha (CT_PositiveFixedPercentage): a color's opacity, in
// thousandths of a percent (0-100000; 100000 is fully opaque).
type Alpha struct {
	XMLName xml.Name `xml:"a:alpha"`
	Val     int      `xml:"val,attr"`
}

// LumMod is a:lumMod (CT_Percentage): scales a color's luminance by the
// given percentage, in thousandths of a percent.
type LumMod struct {
	XMLName xml.Name `xml:"a:lumMod"`
	Val     int      `xml:"val,attr"`
}

// LumOff is a:lumOff (CT_Percentage): shifts a color's luminance by the
// given percentage, in thousandths of a percent — typically paired with
// LumMod to compute the tint/shade variants PowerPoint's own theme-color
// picker offers (e.g. "Accent 1, Lighter 40%").
type LumOff struct {
	XMLName xml.Name `xml:"a:lumOff"`
	Val     int      `xml:"val,attr"`
}

// SysClr (a:sysClr) is a system color: Val names a system color slot (e.g.
// "windowText", "window") whose actual RGB the consumer resolves from the
// viewer's OS at display time, and LastClr is the fallback hex to use when no
// system value is available. Used for a theme's dk1/lt1 slots, where Office
// itself ties the primary text/background to the system colors so a deck
// respects High-Contrast and other OS accessibility settings.
type SysClr struct {
	XMLName xml.Name `xml:"a:sysClr"`
	Val     string   `xml:"val,attr"`
	LastClr string   `xml:"lastClr,attr,omitempty"`
}

// NoFill (a:noFill) is an explicit "no fill" — distinct from omitting a
// fill element, which lets the shape inherit one from its style or layout.
type NoFill struct {
	XMLName xml.Name `xml:"a:noFill"`
}

// GradFill (a:gradFill, CT_GradientFillProperties) is a linear gradient
// fill: an ordered list of color stops (GsLst) blended along an axis (Lin).
// Only the linear-gradient shape (a:lin) is modeled — the schema's other
// gradient path option (a:path, for radial/rectangular/shape gradients) is
// out of scope until a caller needs it. Field order mirrors the schema:
// GsLst before Lin.
type GradFill struct {
	XMLName      xml.Name `xml:"a:gradFill"`
	RotWithShape OnOff    `xml:"rotWithShape,attr,omitempty"`
	GsLst        *GsLst   `xml:"a:gsLst"`
	Lin          *Lin     `xml:"a:lin,omitempty"`
}

// GsLst (a:gsLst, CT_GradientStopList) is a gradient's ordered list of color
// stops. The schema requires at least 2.
type GsLst struct {
	XMLName xml.Name `xml:"a:gsLst"`
	Gs      []*Gs    `xml:"a:gs"`
}

// Gs (a:gs, CT_GradientStop) is one color stop within a gradient: a color
// (as either an explicit RGB value or a theme color reference, the same
// SolidFill choice) at a position along the gradient's axis. Pos is in
// thousandths of a percent (0-100000; e.g. 50000 is the stop's midpoint).
type Gs struct {
	XMLName   xml.Name   `xml:"a:gs"`
	Pos       int        `xml:"pos,attr"`
	SrgbClr   *SrgbClr   `xml:"a:srgbClr,omitempty"`
	SchemeClr *SchemeClr `xml:"a:schemeClr,omitempty"`
}

// Lin (a:lin, CT_LinearShadeProperties) is a linear gradient's direction:
// an angle in 60,000ths of a degree, and whether that angle rotates with
// the shape it fills.
type Lin struct {
	XMLName xml.Name `xml:"a:lin"`
	Ang     int      `xml:"ang,attr"`
	Scaled  OnOff    `xml:"scaled,attr,omitempty"`
}
