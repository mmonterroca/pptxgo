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

// XMLSlideMaster represents ppt/slideMasters/slideMaster1.xml (p:sldMaster).
//
// The walking skeleton models this with structs rather than a hand-typed
// string literal (unlike the theme): cSld/spTree/clrMapOvr already exist as
// typed Go values because slide1.xml needs them, and CT_SlideMaster reuses
// those exact same schema types. Letting the XML encoder produce this part
// removes an entire class of "unclosed tag" or "mismatched attribute" risk
// that a hand-typed literal would carry, for no extra design cost. What
// stays genuinely deferred to a later phase is placeholder inheritance
// (p:ph, idx/type matching across master → layout → slide) — this phase's
// spTree has no shapes at all, just the required-empty group container.
type XMLSlideMaster struct {
	XMLName        xml.Name        `xml:"p:sldMaster"`
	XmlnsA         string          `xml:"xmlns:a,attr"`
	XmlnsR         string          `xml:"xmlns:r,attr"`
	XmlnsP         string          `xml:"xmlns:p,attr"`
	CSld           *CSld           `xml:"p:cSld"`
	ClrMap         *ClrMap         `xml:"p:clrMap"`
	SldLayoutIdLst *SldLayoutIdLst `xml:"p:sldLayoutIdLst"`
	TxStyles       *TxStyles       `xml:"p:txStyles"`
}

// ClrMap is p:clrMap: the required color-slot mapping every slide master
// declares, assigning each of the 12 logical color-map slots to a theme
// scheme color. This is the conventional, near-universal default mapping.
type ClrMap struct {
	XMLName  xml.Name `xml:"p:clrMap"`
	Bg1      string   `xml:"bg1,attr"`
	Tx1      string   `xml:"tx1,attr"`
	Bg2      string   `xml:"bg2,attr"`
	Tx2      string   `xml:"tx2,attr"`
	Accent1  string   `xml:"accent1,attr"`
	Accent2  string   `xml:"accent2,attr"`
	Accent3  string   `xml:"accent3,attr"`
	Accent4  string   `xml:"accent4,attr"`
	Accent5  string   `xml:"accent5,attr"`
	Accent6  string   `xml:"accent6,attr"`
	Hlink    string   `xml:"hlink,attr"`
	FolHlink string   `xml:"folHlink,attr"`
}

// NewDefaultClrMap returns the standard bg/tx-to-theme-slot mapping used by
// virtually every OOXML presentation.
func NewDefaultClrMap() *ClrMap {
	return &ClrMap{
		Bg1: "lt1", Tx1: "dk1", Bg2: "lt2", Tx2: "dk2",
		Accent1: "accent1", Accent2: "accent2", Accent3: "accent3",
		Accent4: "accent4", Accent5: "accent5", Accent6: "accent6",
		Hlink: "hlink", FolHlink: "folHlink",
	}
}

// SldLayoutIdLst is p:sldLayoutIdLst: the list of layouts owned by a master.
type SldLayoutIdLst struct {
	XMLName xml.Name       `xml:"p:sldLayoutIdLst"`
	Entries []*SldLayoutId `xml:"p:sldLayoutId"`
}

// SldLayoutId is a single p:sldLayoutId entry.
type SldLayoutId struct {
	XMLName xml.Name `xml:"p:sldLayoutId"`
	ID      uint32   `xml:"id,attr"`
	RID     string   `xml:"r:id,attr"`
}

// TxStyles is p:txStyles: default text formatting for title, body, and
// other placeholders, cascaded down to every layout and slide that doesn't
// override it. A single first-level definition per style is schema-valid;
// PowerPoint itself writes nine cascading levels, but nothing requires it.
type TxStyles struct {
	XMLName    xml.Name   `xml:"p:txStyles"`
	TitleStyle *TextStyle `xml:"p:titleStyle"`
	BodyStyle  *TextStyle `xml:"p:bodyStyle"`
	OtherStyle *TextStyle `xml:"p:otherStyle"`
}

// TextStyle wraps a single cascading paragraph-level style definition. It
// deliberately has no XMLName field: TxStyles reuses this one type for
// p:titleStyle, p:bodyStyle, and p:otherStyle, and an XMLName field on the
// child would override the parent field's tag and force the same element
// name onto all three.
type TextStyle struct {
	Lvl1PPr *Lvl1PPr `xml:"a:lvl1pPr"`
}

// Lvl1PPr is a:lvl1pPr: the first (and, here, only) indentation level's
// paragraph properties — the same content model as drawingml.PPr (both are
// CT_TextParagraphProperties), but modeled as its own type: PPr's own
// fixed XMLName ("a:pPr") would win over TxStyles' "a:lvl1pPr" field tag if
// embedded directly, the same reuse trap Xfrm/GraphicFrameXfrm document,
// and unlike a:pPr this element also carries a trailing a:defRPr with the
// level's default run properties. Field order mirrors the schema: MarL/
// Indent attrs, then BuFont ahead of the mutually-exclusive bullet group
// (BuNone/BuChar), then DefRPr last.
type Lvl1PPr struct {
	XMLName xml.Name          `xml:"a:lvl1pPr"`
	MarL    *int              `xml:"marL,attr,omitempty"`
	Indent  *int              `xml:"indent,attr,omitempty"`
	BuFont  *drawingml.BuFont `xml:"a:buFont,omitempty"`
	BuNone  *drawingml.BuNone `xml:"a:buNone,omitempty"`
	BuChar  *drawingml.BuChar `xml:"a:buChar,omitempty"`
	DefRPr  *DefRPr           `xml:"a:defRPr"`
}

// DefRPr is a:defRPr: default run (character) properties for a paragraph level.
type DefRPr struct {
	XMLName xml.Name `xml:"a:defRPr"`
	Sz      int      `xml:"sz,attr,omitempty"` // hundredths of a point
}

// bodyBulletMarL and bodyBulletIndent are the conventional level-1 hanging
// indent for a bulleted body placeholder, in EMUs (0.375in each): the
// bullet glyph sits at MarL+Indent (0) and wrapped text at MarL.
const (
	bodyBulletMarL   = 342900
	bodyBulletIndent = -342900
)

// NewDefaultTxStyles returns a minimal title/body/other text style set with
// conventional default sizes (44pt title, 32pt body, 18pt other). The body
// style also carries a level-1 bullet default (a round "•" in Arial) so a
// body placeholder's paragraphs pick up a bullet automatically unless they
// set their own (Paragraph.Bullet/NumberedBullet) or explicitly suppress it
// (Paragraph.NoBullet) — pptxgo's txBody always emits its own a:lstStyle
// empty, so nothing on the placeholder itself overrides this cascade.
func NewDefaultTxStyles() *TxStyles {
	marL, indent := bodyBulletMarL, bodyBulletIndent
	return &TxStyles{
		TitleStyle: &TextStyle{Lvl1PPr: &Lvl1PPr{DefRPr: &DefRPr{Sz: 4400}}},
		BodyStyle: &TextStyle{Lvl1PPr: &Lvl1PPr{
			MarL:   &marL,
			Indent: &indent,
			BuFont: &drawingml.BuFont{Typeface: "Arial"},
			BuChar: &drawingml.BuChar{Char: "•"},
			DefRPr: &DefRPr{Sz: 3200},
		}},
		OtherStyle: &TextStyle{Lvl1PPr: &Lvl1PPr{DefRPr: &DefRPr{Sz: 1800}}},
	}
}

// XMLSlideLayout represents ppt/slideLayouts/slideLayout1.xml (p:sldLayout).
type XMLSlideLayout struct {
	XMLName   xml.Name   `xml:"p:sldLayout"`
	XmlnsA    string     `xml:"xmlns:a,attr"`
	XmlnsR    string     `xml:"xmlns:r,attr"`
	XmlnsP    string     `xml:"xmlns:p,attr"`
	Type      string     `xml:"type,attr,omitempty"`
	CSld      *CSld      `xml:"p:cSld"`
	ClrMapOvr *ClrMapOvr `xml:"p:clrMapOvr"`
}
