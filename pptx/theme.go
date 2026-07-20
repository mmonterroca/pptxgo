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
	"bytes"
	"encoding/xml"

	"github.com/mmonterroca/pptxgo/drawingml"
)

// Theme is a presentation's visual identity: the color scheme and font
// scheme (ppt/theme/theme1.xml's a:clrScheme and a:fontScheme) that every
// slide inherits. Pass one to New via WithTheme to brand a whole deck at
// once — because every shape/text/background can reference a theme color by
// slot (see SchemeColor, FillScheme, ColorScheme, BackgroundScheme) rather
// than a hardcoded RGB, swapping the Theme recolors all of them with no
// call-site changes.
//
// Only the brand-relevant parts of a theme are modeled: the twelve color
// slots and the two font typefaces. The format scheme (a:fmtScheme — the
// fill/line/effect style *definitions* PowerPoint's own themes carry) is
// kept at Office's standard values, since a brand deck varies its palette
// and typography, not those low-level style-list definitions.
type Theme struct {
	// Name is the theme's display name (a:theme/@name and the color/font
	// scheme names). Empty defaults to "Office".
	Name string

	// Colors is the twelve-slot color scheme (a:clrScheme).
	Colors ThemeColors

	// Fonts is the major/minor font scheme (a:fontScheme).
	Fonts ThemeFonts
}

// ThemeColors is a theme's twelve-slot color scheme (a:clrScheme). Dark1/
// Light1 are the primary text/background pair (conventionally near-black and
// near-white); Dark2/Light2 the secondary pair; Accent1-6 the accent palette;
// Hyperlink/FollowedHyperlink the two link colors. A slide references these
// through its color map (see NewDefaultClrMap) — e.g. SchemeAccent1 resolves
// to Accent1, SchemeText1/SchemeBackground1 to Dark1/Light1.
type ThemeColors struct {
	Dark1             drawingml.Color
	Light1            drawingml.Color
	Dark2             drawingml.Color
	Light2            drawingml.Color
	Accent1           drawingml.Color
	Accent2           drawingml.Color
	Accent3           drawingml.Color
	Accent4           drawingml.Color
	Accent5           drawingml.Color
	Accent6           drawingml.Color
	Hyperlink         drawingml.Color
	FollowedHyperlink drawingml.Color
}

// ThemeFonts is a theme's font scheme (a:fontScheme): the major (heading)
// and minor (body) Latin typefaces. Placeholder and default text with no
// explicit Font inherits the minor font; PowerPoint's own "+headings"/"+body"
// font choices resolve to these. An empty typeface defaults to Office's
// (Calibri Light major, Calibri minor).
type ThemeFonts struct {
	Major string // headings — a:majorFont's Latin typeface (e.g. "Calibri Light")
	Minor string // body — a:minorFont's Latin typeface (e.g. "Calibri")
}

// DefaultTheme returns Office's standard theme — the palette and typography
// New uses when no WithTheme option is given. Start from it to tweak only a
// few slots:
//
//	t := pptx.DefaultTheme()
//	t.Name = "Acme"
//	t.Colors.Accent1 = pptx.RGB(0x1F, 0x49, 0x7D)
//	p := pptx.New(pptx.WithTheme(t))
func DefaultTheme() Theme {
	return Theme{
		Name: "Office",
		Colors: ThemeColors{
			Dark1:             drawingml.Color{R: 0x00, G: 0x00, B: 0x00},
			Light1:            drawingml.Color{R: 0xFF, G: 0xFF, B: 0xFF},
			Dark2:             drawingml.Color{R: 0x44, G: 0x54, B: 0x6A},
			Light2:            drawingml.Color{R: 0xE7, G: 0xE6, B: 0xE6},
			Accent1:           drawingml.Color{R: 0x44, G: 0x72, B: 0xC4},
			Accent2:           drawingml.Color{R: 0xED, G: 0x7D, B: 0x31},
			Accent3:           drawingml.Color{R: 0xA5, G: 0xA5, B: 0xA5},
			Accent4:           drawingml.Color{R: 0xFF, G: 0xC0, B: 0x00},
			Accent5:           drawingml.Color{R: 0x5B, G: 0x9B, B: 0xD5},
			Accent6:           drawingml.Color{R: 0x70, G: 0xAD, B: 0x47},
			Hyperlink:         drawingml.Color{R: 0x05, G: 0x63, B: 0xC1},
			FollowedHyperlink: drawingml.Color{R: 0x95, G: 0x4F, B: 0x72},
		},
		Fonts: ThemeFonts{
			Major: "Calibri Light",
			Minor: "Calibri",
		},
	}
}

// themeName returns the theme's display name, defaulting to "Office".
func (t Theme) themeName() string {
	if t.Name == "" {
		return "Office"
	}
	return t.Name
}

// withDefaults returns a copy of t with any unset field filled from
// DefaultTheme: an empty font typeface, or a color slot left at its zero value
// (the zero drawingml.Color, pure black), falls back to the default's value.
// This makes a partially-specified Theme safe — passing
// Theme{Colors: ThemeColors{Accent1: brandNavy}} to WithTheme keeps the other
// eleven slots at the default Office palette instead of rendering them as
// black — mirroring the empty-font fallback. The one edge case: a slot
// deliberately set to pure black RGB(0,0,0) reads as unset and so inherits the
// default; use the near-black dark slots, or start from DefaultTheme() (whose
// Dark1 is already black), when you need an explicit black.
func (t Theme) withDefaults() Theme {
	d := DefaultTheme()
	fill := func(c *drawingml.Color, def drawingml.Color) {
		if *c == (drawingml.Color{}) {
			*c = def
		}
	}
	fill(&t.Colors.Dark1, d.Colors.Dark1)
	fill(&t.Colors.Light1, d.Colors.Light1)
	fill(&t.Colors.Dark2, d.Colors.Dark2)
	fill(&t.Colors.Light2, d.Colors.Light2)
	fill(&t.Colors.Accent1, d.Colors.Accent1)
	fill(&t.Colors.Accent2, d.Colors.Accent2)
	fill(&t.Colors.Accent3, d.Colors.Accent3)
	fill(&t.Colors.Accent4, d.Colors.Accent4)
	fill(&t.Colors.Accent5, d.Colors.Accent5)
	fill(&t.Colors.Accent6, d.Colors.Accent6)
	fill(&t.Colors.Hyperlink, d.Colors.Hyperlink)
	fill(&t.Colors.FollowedHyperlink, d.Colors.FollowedHyperlink)
	if t.Fonts.Major == "" {
		t.Fonts.Major = d.Fonts.Major
	}
	if t.Fonts.Minor == "" {
		t.Fonts.Minor = d.Fonts.Minor
	}
	return t
}

// clrSlot is one color-scheme slot (a:dk1, a:accent1, ...). The wrapping
// element name comes from the parent field tag in clrSchemeXML; the child is
// exactly one of an explicit a:srgbClr (accents, dk2/lt2, links) or a system
// a:sysClr (dk1/lt1, so the primary text/background follow the OS — see
// SysClr).
type clrSlot struct {
	Srgb *drawingml.SrgbClr `xml:"a:srgbClr,omitempty"`
	Sys  *drawingml.SysClr  `xml:"a:sysClr,omitempty"`
}

func newClrSlot(c drawingml.Color) clrSlot {
	return clrSlot{Srgb: &drawingml.SrgbClr{Val: drawingml.ToHex(c)}}
}

// newSysClrSlot builds a dk1/lt1 slot as a system color (val "windowText" or
// "window") whose lastClr fallback is the theme's own hex — matching how
// Office authors dk1/lt1 so the deck respects OS accessibility settings while
// still carrying a concrete fallback color.
func newSysClrSlot(val string, c drawingml.Color) clrSlot {
	return clrSlot{Sys: &drawingml.SysClr{Val: val, LastClr: drawingml.ToHex(c)}}
}

// clrSchemeXML models a:clrScheme so the twelve brand color slots are
// encoder-generated rather than hand-spliced into the theme string.
type clrSchemeXML struct {
	XMLName  xml.Name `xml:"a:clrScheme"`
	Name     string   `xml:"name,attr"`
	Dk1      clrSlot  `xml:"a:dk1"`
	Lt1      clrSlot  `xml:"a:lt1"`
	Dk2      clrSlot  `xml:"a:dk2"`
	Lt2      clrSlot  `xml:"a:lt2"`
	Accent1  clrSlot  `xml:"a:accent1"`
	Accent2  clrSlot  `xml:"a:accent2"`
	Accent3  clrSlot  `xml:"a:accent3"`
	Accent4  clrSlot  `xml:"a:accent4"`
	Accent5  clrSlot  `xml:"a:accent5"`
	Accent6  clrSlot  `xml:"a:accent6"`
	Hlink    clrSlot  `xml:"a:hlink"`
	FolHlink clrSlot  `xml:"a:folHlink"`
}

// latinFontXML is one script slot (a:latin/a:ea/a:cs) inside a font
// collection. typeface is always emitted (empty for ea/cs, matching Office).
type latinFontXML struct {
	Typeface string `xml:"typeface,attr"`
}

// fontCollectionXML models a:majorFont / a:minorFont — its element name comes
// from the parent field tag in fontSchemeXML.
type fontCollectionXML struct {
	Latin latinFontXML `xml:"a:latin"`
	Ea    latinFontXML `xml:"a:ea"`
	Cs    latinFontXML `xml:"a:cs"`
}

// fontSchemeXML models a:fontScheme so the brand typefaces are
// encoder-generated.
type fontSchemeXML struct {
	XMLName   xml.Name          `xml:"a:fontScheme"`
	Name      string            `xml:"name,attr"`
	MajorFont fontCollectionXML `xml:"a:majorFont"`
	MinorFont fontCollectionXML `xml:"a:minorFont"`
}

// themeFmtScheme is a:fmtScheme, the fill/line/effect/background style
// definitions PowerPoint themes carry. It is held as a fixed, valid Office
// block because a brand deck varies its palette and typography (both
// modeled above), not these low-level style-list definitions — modeling
// them as structs would buy nothing until something needs to vary them, the
// same reasoning the whole theme was a literal under before WithTheme.
const themeFmtScheme = `<a:fmtScheme name="Office">` +
	`<a:fillStyleLst>` +
	`<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
	`<a:gradFill rotWithShape="1"><a:gsLst>` +
	`<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>` +
	`<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>` +
	`<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>` +
	`</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>` +
	`<a:gradFill rotWithShape="1"><a:gsLst>` +
	`<a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>` +
	`<a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>` +
	`<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="350000"/></a:schemeClr></a:gs>` +
	`</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>` +
	`</a:fillStyleLst>` +
	`<a:lnStyleLst>` +
	`<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>` +
	`<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>` +
	`<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>` +
	`</a:lnStyleLst>` +
	`<a:effectStyleLst>` +
	`<a:effectStyle><a:effectLst/></a:effectStyle>` +
	`<a:effectStyle><a:effectLst/></a:effectStyle>` +
	`<a:effectStyle><a:effectLst>` +
	`<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw>` +
	`</a:effectLst></a:effectStyle>` +
	`</a:effectStyleLst>` +
	`<a:bgFillStyleLst>` +
	`<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
	`<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>` +
	`<a:gradFill rotWithShape="1"><a:gsLst>` +
	`<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>` +
	`<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>` +
	`<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>` +
	`</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>` +
	`</a:bgFillStyleLst>` +
	`</a:fmtScheme>`

// renderThemeXML produces the ppt/theme/theme1.xml bytes for t. The varying
// parts (color scheme, font scheme) are encoder-generated from structs — no
// hand-spliced tags — and wrapped in the fixed, valid a:theme scaffold that
// declares the a:namespace those fragments' "a:"-prefixed tags rely on.
func renderThemeXML(t Theme) []byte {
	name := t.themeName()
	// Fill any unset color slot / empty font from the default palette, so a
	// partially-specified Theme doesn't render its untouched slots as black.
	t = t.withDefaults()

	clr, err := xml.Marshal(clrSchemeXML{
		Name: name,
		// dk1/lt1 are system colors (Office convention) so the primary
		// text/background follow the viewer's OS accessibility settings, with
		// the theme's own hex as the lastClr fallback.
		Dk1:      newSysClrSlot("windowText", t.Colors.Dark1),
		Lt1:      newSysClrSlot("window", t.Colors.Light1),
		Dk2:      newClrSlot(t.Colors.Dark2),
		Lt2:      newClrSlot(t.Colors.Light2),
		Accent1:  newClrSlot(t.Colors.Accent1),
		Accent2:  newClrSlot(t.Colors.Accent2),
		Accent3:  newClrSlot(t.Colors.Accent3),
		Accent4:  newClrSlot(t.Colors.Accent4),
		Accent5:  newClrSlot(t.Colors.Accent5),
		Accent6:  newClrSlot(t.Colors.Accent6),
		Hlink:    newClrSlot(t.Colors.Hyperlink),
		FolHlink: newClrSlot(t.Colors.FollowedHyperlink),
	})
	if err != nil {
		panic(err) // static struct with string/hex fields only; cannot fail
	}

	font, err := xml.Marshal(fontSchemeXML{
		Name:      name,
		MajorFont: fontCollectionXML{Latin: latinFontXML{Typeface: t.Fonts.Major}},
		MinorFont: fontCollectionXML{Latin: latinFontXML{Typeface: t.Fonts.Minor}},
	})
	if err != nil {
		panic(err) // static struct with string fields only; cannot fail
	}

	var b bytes.Buffer
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	b.WriteString(`<a:theme xmlns:a="`)
	b.WriteString(drawingml.NamespaceMain)
	b.WriteString(`" name="`)
	// EscapeText escapes the reserved characters (&, <, >, ", ') so a brand
	// name is safe as an attribute value.
	if err := xml.EscapeText(&b, []byte(name)); err != nil {
		panic(err)
	}
	b.WriteString(`"><a:themeElements>`)
	b.Write(clr)
	b.Write(font)
	b.WriteString(themeFmtScheme)
	b.WriteString(`</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`)
	return b.Bytes()
}
