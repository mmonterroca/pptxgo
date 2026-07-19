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

// TextStyle wraps a cascading paragraph-level style definition, one LvlPPr
// per indentation level it defines (1-9; Paragraph.Level(0) through
// Level(8) select among them, 0-indexed there vs. 1-indexed in the OOXML
// element names). It deliberately has no XMLName field and instead
// implements MarshalXML: TxStyles reuses this one type for p:titleStyle,
// p:bodyStyle, and p:otherStyle (an XMLName field on the child would
// override the parent field's tag and force the same element name onto all
// three, the same reuse trap LvlPPr documents one level down), and each
// entry in Levels needs ITS OWN element name (a:lvl1pPr, a:lvl2pPr, ...) —
// something a single static field tag can't express either.
type TextStyle struct {
	Levels []*LvlPPr
}

// MarshalXML emits start (whatever name the parent's own field tag gave
// this TextStyle — see the type doc) followed by each of Levels, each
// self-naming via LvlPPr.MarshalXML.
func (ts *TextStyle) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	for _, lvl := range ts.Levels {
		if err := e.Encode(lvl); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// LvlPPr is a:lvl1pPr through a:lvl9pPr (selected by Level, 1-9) — the same
// content model as drawingml.PPr (both are CT_TextParagraphProperties), but
// modeled as its own type: PPr's own fixed XMLName ("a:pPr") would win over
// a field tag if embedded directly, the same reuse trap TextStyle documents
// one level up, and unlike a:pPr this element also carries a trailing
// a:defRPr with the level's default run properties. A single type serves
// all nine levels (rather than nine near-identical structs) via
// MarshalXML choosing the element name from Level; field order within it
// mirrors the schema: MarL/Indent attrs, then BuFont ahead of the
// mutually-exclusive bullet group (BuNone/BuChar), then DefRPr last.
type LvlPPr struct {
	Level  int // 1-9; selects this level's element name — see MarshalXML
	MarL   *int
	Indent *int
	BuFont *drawingml.BuFont
	BuNone *drawingml.BuNone
	BuChar *drawingml.BuChar
	DefRPr *DefRPr
}

// lvlPPrContent mirrors LvlPPr's field set (minus Level, which selects the
// element name rather than being marshaled itself) under fixed tags, so
// MarshalXML can delegate the actual attribute/child encoding to the
// standard encoding/xml struct-tag walk instead of hand-writing token
// output.
type lvlPPrContent struct {
	MarL   *int              `xml:"marL,attr,omitempty"`
	Indent *int              `xml:"indent,attr,omitempty"`
	BuFont *drawingml.BuFont `xml:"a:buFont,omitempty"`
	BuNone *drawingml.BuNone `xml:"a:buNone,omitempty"`
	BuChar *drawingml.BuChar `xml:"a:buChar,omitempty"`
	DefRPr *DefRPr           `xml:"a:defRPr"`
}

// MarshalXML implements xml.Marshaler, naming the element a:lvl<Level>pPr.
// It rejects a Level outside 1-9 with an error rather than emitting an
// out-of-schema element name: the old fixed-XMLName Lvl1PPr made an invalid
// name structurally impossible, and since LvlPPr/TextStyle.Levels are
// exported (a caller can build a &LvlPPr{} directly, e.g. omitting Level so
// it defaults to 0, or setting 10+), that guarantee is re-established here.
// The error surfaces at Save time — pptxgo's own newBodyLevels/
// NewDefaultTxStyles always set Level in range, so this only fires on
// caller-constructed styles.
func (l *LvlPPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if l.Level < 1 || l.Level > 9 {
		return errors.InvalidArgument("LvlPPr.MarshalXML", "Level", l.Level, "must be between 1 and 9 (a:lvl1pPr through a:lvl9pPr)")
	}
	start.Name = xml.Name{Local: fmt.Sprintf("a:lvl%dpPr", l.Level)}
	return e.EncodeElement(&lvlPPrContent{
		MarL: l.MarL, Indent: l.Indent,
		BuFont: l.BuFont, BuNone: l.BuNone, BuChar: l.BuChar,
		DefRPr: l.DefRPr,
	}, start)
}

// DefRPr is a:defRPr: default run (character) properties for a paragraph level.
type DefRPr struct {
	XMLName xml.Name `xml:"a:defRPr"`
	Sz      int      `xml:"sz,attr,omitempty"` // hundredths of a point
}

// bodyBulletMarL and bodyBulletIndent are the conventional level-1 hanging
// indent for a bulleted body placeholder, in EMUs (0.375in each): the
// bullet glyph sits at MarL+Indent (0) and wrapped text at MarL. Deeper
// levels (see newBodyLevels) grow MarL by this same increment per level,
// keeping Indent constant — only the base indent grows, not the hanging
// gap between the bullet glyph and its text.
const (
	bodyBulletMarL   = 342900
	bodyBulletIndent = -342900
)

// bodyLevelCount is how many of the schema's 9 available levels
// newBodyLevels populates — every level Paragraph.Level(0..8) can select.
const bodyLevelCount = 9

// bodyBulletGlyphs cycles the bullet character used at each body level,
// alternating "•"/"–" the way PowerPoint's own built-in themes do, so nine
// levels of nested bullets stay visually distinguishable from their parent.
var bodyBulletGlyphs = [bodyLevelCount]string{"•", "–", "•", "–", "•", "–", "•", "–", "•"}

// bodyLevelSizes is the default font size (hundredths of a point) per body
// level. Deeper levels shrink and then plateau, mirroring how PowerPoint's
// own built-in themes de-emphasize nested content — a flat size across all
// nine levels renders a Level(8) sub-bullet as large as a top-level one,
// defeating the visual hierarchy the indent and alternating glyphs
// establish. Level 1 stays 32pt (the body style's headline default the
// title/other styles are sized against).
var bodyLevelSizes = [bodyLevelCount]int{3200, 2800, 2400, 2000, 2000, 1800, 1800, 1800, 1800}

// newBodyLevels returns bodyLevelCount LvlPPr entries (levels 1-9) for the
// body style's bullet cascade — see NewDefaultTxStyles.
func newBodyLevels() []*LvlPPr {
	levels := make([]*LvlPPr, bodyLevelCount)
	for i := range levels {
		marL := bodyBulletMarL * (i + 1)
		indent := bodyBulletIndent
		levels[i] = &LvlPPr{
			Level:  i + 1,
			MarL:   &marL,
			Indent: &indent,
			BuFont: &drawingml.BuFont{Typeface: "Arial"},
			BuChar: &drawingml.BuChar{Char: bodyBulletGlyphs[i]},
			DefRPr: &DefRPr{Sz: bodyLevelSizes[i]},
		}
	}
	return levels
}

// NewDefaultTxStyles returns a minimal title/body/other text style set with
// conventional default sizes (44pt title, 32pt body, 18pt other). The body
// style carries a bullet default (alternating "•"/"–" in Arial) across all
// 9 levels so a body placeholder's paragraphs — at any Paragraph.Level(0..8)
// — pick up a bullet and indent automatically unless they set their own
// (Paragraph.Bullet/NumberedBullet) or explicitly suppress it
// (Paragraph.NoBullet) — pptxgo's txBody always emits its own a:lstStyle
// empty, so nothing on the placeholder itself overrides this cascade.
// TitleStyle/OtherStyle keep just their first level: pptxgo never applies a
// Level to a title or "other" placeholder's paragraphs, so levels 2-9
// would go unused there.
func NewDefaultTxStyles() *TxStyles {
	return &TxStyles{
		TitleStyle: &TextStyle{Levels: []*LvlPPr{{Level: 1, DefRPr: &DefRPr{Sz: 4400}}}},
		BodyStyle:  &TextStyle{Levels: newBodyLevels()},
		OtherStyle: &TextStyle{Levels: []*LvlPPr{{Level: 1, DefRPr: &DefRPr{Sz: 1800}}}},
	}
}

// XMLSlideLayout represents one ppt/slideLayouts/slideLayoutN.xml part
// (p:sldLayout) — see SlideLayoutPath and newStandardLayouts.
type XMLSlideLayout struct {
	XMLName   xml.Name   `xml:"p:sldLayout"`
	XmlnsA    string     `xml:"xmlns:a,attr"`
	XmlnsR    string     `xml:"xmlns:r,attr"`
	XmlnsP    string     `xml:"xmlns:p,attr"`
	Type      string     `xml:"type,attr,omitempty"`
	CSld      *CSld      `xml:"p:cSld"`
	ClrMapOvr *ClrMapOvr `xml:"p:clrMapOvr"`
}
