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

// EffectLst is a:effectLst (CT_EffectList): a shape or picture's visual
// effects. Blur, fill-overlay, inner-shadow, and preset-shadow are out of
// scope until a caller needs them — outer shadow, glow, reflection, and
// soft edge are modeled. Unlike SrgbClr/SchemeClr's color transforms, this
// is a true xsd:sequence (each effect optional, but in a fixed relative
// order when present), so field order here IS schema-mandated: glow,
// outerShdw, reflection, softEdge.
type EffectLst struct {
	XMLName    xml.Name    `xml:"a:effectLst"`
	Glow       *Glow       `xml:"a:glow,omitempty"`
	OuterShdw  *OuterShdw  `xml:"a:outerShdw,omitempty"`
	Reflection *Reflection `xml:"a:reflection,omitempty"`
	SoftEdge   *SoftEdge   `xml:"a:softEdge,omitempty"`
}

// Glow is a:glow (CT_GlowEffect): a soft-edged color halo drawn around a
// shape's own outline. Rad is the glow's radius, in EMUs.
type Glow struct {
	XMLName   xml.Name   `xml:"a:glow"`
	Rad       int        `xml:"rad,attr,omitempty"`
	SrgbClr   *SrgbClr   `xml:"a:srgbClr,omitempty"`
	SchemeClr *SchemeClr `xml:"a:schemeClr,omitempty"`
}

// OuterShdw is a:outerShdw (CT_OuterShadowEffect): a drop shadow cast
// outside the shape. BlurRad and Dist are in EMUs, Dir in 60,000ths of a
// degree. RotWithShape is a *TriState (not OnOff) because Office's own
// default outer-shadow preset explicitly writes rotWithShape="0" — the
// schema's own default is true, so omitting it here would silently give the
// opposite (shadow rotates with the shape) behavior once the shape itself
// is rotated. Scale/skew (sx/sy/kx/ky) are out of scope until a caller
// needs them.
type OuterShdw struct {
	XMLName      xml.Name   `xml:"a:outerShdw"`
	BlurRad      int        `xml:"blurRad,attr,omitempty"`
	Dist         int        `xml:"dist,attr,omitempty"`
	Dir          int        `xml:"dir,attr,omitempty"`
	Algn         string     `xml:"algn,attr,omitempty"`
	RotWithShape *TriState  `xml:"rotWithShape,attr,omitempty"`
	SrgbClr      *SrgbClr   `xml:"a:srgbClr,omitempty"`
	SchemeClr    *SchemeClr `xml:"a:schemeClr,omitempty"`
}

// Reflection is a:reflection (CT_ReflectionEffect): a mirror-image
// reflection of the shape, fading in opacity along its own axis. Only the
// fade/blur/direction attributes needed for a straight-down reflection are
// modeled — scale/skew (sx/sy/kx/ky), FadeDir, and Algn are out of scope
// until a caller needs them.
type Reflection struct {
	XMLName xml.Name `xml:"a:reflection"`
	BlurRad int      `xml:"blurRad,attr,omitempty"`
	StA     int      `xml:"stA,attr,omitempty"`
	StPos   int      `xml:"stPos,attr,omitempty"`
	EndA    int      `xml:"endA,attr,omitempty"`
	EndPos  int      `xml:"endPos,attr,omitempty"`
	Dist    int      `xml:"dist,attr,omitempty"`
	Dir     int      `xml:"dir,attr,omitempty"`
}

// SoftEdge is a:softEdge (CT_SoftEdgesEffect): fades the shape's own edges
// to transparent over the given radius, in EMUs.
type SoftEdge struct {
	XMLName xml.Name `xml:"a:softEdge"`
	Rad     int      `xml:"rad,attr"`
}

// TriState models an xsd:boolean attribute where an explicit "false" must
// be distinguishable from "not set" — unlike OnOff, which can only ever
// represent true or omitted (see OnOff's own doc comment), a *TriState
// field can also emit an explicit "0". A nil field omits the attribute
// entirely, letting the schema's own default apply; a non-nil value
// marshals as Office's own "1"/"0" convention rather than encoding/xml's
// default "true"/"false".
type TriState bool

// MarshalXMLAttr implements xml.MarshalerAttr.
func (t TriState) MarshalXMLAttr(name xml.Name) (xml.Attr, error) {
	if t {
		return xml.Attr{Name: name, Value: "1"}, nil
	}
	return xml.Attr{Name: name, Value: "0"}, nil
}
