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

// TextBody is the DrawingML text-body model (CT_TextBody): body-level
// formatting properties, an optional list-style, and one or more
// paragraphs. Field order mirrors the schema's xsd:sequence — BodyPr, then
// LstStyle, then Paragraphs — because the OpenXML SDK validator rejects
// children emitted out of order, a defect no strings.Contains test catches.
//
// TextBody deliberately has no XMLName field. In PresentationML the element
// is p:txBody (namespace p, children a:); a fixed XMLName here would win
// over the host's field tag (encoding/xml's rule: a type's own XMLName beats
// the tag on the field that embeds it) and hardcode a DOCX- or PPTX-specific
// prefix into a package that must stay agnostic of both. The containing
// element name is left entirely to the field tag the host package chooses
// (e.g. pptx.Shape.TxBody `xml:"p:txBody"`).
type TextBody struct {
	BodyPr     *BodyPr      `xml:"a:bodyPr"`
	LstStyle   *LstStyle    `xml:"a:lstStyle,omitempty"`
	Paragraphs []*Paragraph `xml:"a:p"`
}

// MarshalXML fills two schema minimums a caller can otherwise leave unmet
// by constructing a TextBody directly (bypassing pptx.Slide.AddTextBox,
// which always sets both): CT_TextBody requires a:bodyPr before anything
// else, and its p child has minOccurs="1". A bare &TextBody{} therefore
// still emits a valid <a:bodyPr/> and a single empty <a:p/> rather than
// schema-invalid missing/empty content. A type alias breaks the recursion
// that would otherwise result from calling back into this same method.
func (b *TextBody) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	type alias TextBody
	out := *b
	if out.BodyPr == nil {
		out.BodyPr = &BodyPr{}
	}
	if len(out.Paragraphs) == 0 {
		out.Paragraphs = []*Paragraph{{}}
	}
	return e.EncodeElement((*alias)(&out), start)
}

// BodyPr is a:bodyPr (CT_TextBodyProperties): body-level text formatting.
// Attribute order doesn't matter (XML attributes are unordered), but the
// autofit choice (NoAutofit/NormAutofit/SpAutoFit — CT_TextBodyProperties'
// EG_TextAutofit group) is schema-required to come first among children;
// since no other child element is modeled, that group is also the only one
// emitted, so there is nothing to order it against. The zero value marshals
// as the minimal valid <a:bodyPr/>.
type BodyPr struct {
	XMLName xml.Name `xml:"a:bodyPr"`
	Wrap    string   `xml:"wrap,attr,omitempty"` // ST_TextWrappingType: "none" or "square"
	// LIns/TIns/RIns/BIns are *int, not int, so an explicit 0 (e.g. a
	// caller removing the default inset entirely) still marshals as
	// lIns="0" instead of being indistinguishable from "never set" and
	// silently dropped by omitempty — PowerPoint's own default insets are
	// non-zero, so a dropped explicit 0 changes the rendered layout.
	LIns        *int         `xml:"lIns,attr,omitempty"`   // left inset, EMUs
	TIns        *int         `xml:"tIns,attr,omitempty"`   // top inset, EMUs
	RIns        *int         `xml:"rIns,attr,omitempty"`   // right inset, EMUs
	BIns        *int         `xml:"bIns,attr,omitempty"`   // bottom inset, EMUs
	Anchor      string       `xml:"anchor,attr,omitempty"` // ST_TextAnchoringType: "t", "ctr", "b"
	NoAutofit   *NoAutofit   `xml:"a:noAutofit,omitempty"`
	NormAutofit *NormAutofit `xml:"a:normAutofit,omitempty"`
	SpAutoFit   *SpAutoFit   `xml:"a:spAutoFit,omitempty"`
}

// NoAutofit is a:noAutofit (CT_TextNoAutofit): text may overflow the shape
// uncorrected. Emitted empty — PowerPoint does not parameterize this choice.
type NoAutofit struct {
	XMLName xml.Name `xml:"a:noAutofit"`
}

// NormAutofit is a:normAutofit (CT_TextNormalAutofit): shrink font size
// and/or line spacing to fit the text within the shape. Emitted empty — its
// optional fontScale/lnSpcReduction attributes are the values PowerPoint
// itself computes on reflow, not something a caller sets up front.
type NormAutofit struct {
	XMLName xml.Name `xml:"a:normAutofit"`
}

// SpAutoFit is a:spAutoFit (CT_TextShapeAutofit): resize the shape to fit
// its text. Emitted empty; the schema defines no attributes for it.
type SpAutoFit struct {
	XMLName xml.Name `xml:"a:spAutoFit"`
}

// LstStyle is a:lstStyle: list-level style overrides. Emitted empty, per
// convention — pptxgo does not yet model per-level overrides.
type LstStyle struct {
	XMLName xml.Name `xml:"a:lstStyle"`
}

// Paragraph is a:p (CT_TextParagraph): paragraph-level properties followed
// by its content. Content is left as `any` — like SpTree.Content in the
// pptx package — because CT_TextParagraph's body is a mixed sequence of
// r|br|fld elements (runs and explicit line breaks, interleaved in
// caller-chosen order), which a single typed slice can't represent; each
// element (*Run, *Br, ...) supplies its own XMLName so `xml:",any"`
// marshals it correctly regardless of position.
type Paragraph struct {
	XMLName xml.Name `xml:"a:p"`
	PPr     *PPr     `xml:"a:pPr,omitempty"`
	Content []any    `xml:",any"`
}

// PPr is a:pPr (CT_TextParagraphProperties): paragraph properties. Field
// order mirrors the schema: attributes (marL, lvl, indent, algn) first,
// then the child sequence — LnSpc, SpcBef, SpcAft, then the bullet-typeface
// group (BuFont) ahead of the mutually-exclusive bullet group (BuNone,
// BuAutoNum, BuChar) — the validator rejects e.g. a:buChar emitted before
// a:buFont. At most one of BuNone/BuAutoNum/BuChar should be set; the
// Paragraph builder methods (NoBullet, NumberedBullet, Bullet) enforce that
// by clearing the other two whenever one is set.
type PPr struct {
	XMLName xml.Name `xml:"a:pPr"`
	// MarL/Indent are *int, not int, for the same reason BodyPr's insets
	// are: an explicit 0 (e.g. overriding an inherited non-zero margin)
	// must marshal as marL="0", not be dropped by omitempty as if it had
	// never been set.
	MarL      *int         `xml:"marL,attr,omitempty"`   // left margin, EMUs
	Lvl       int          `xml:"lvl,attr,omitempty"`    // outline level, 0-8
	Indent    *int         `xml:"indent,attr,omitempty"` // first-line indent, EMUs; negative for hanging indent
	Algn      string       `xml:"algn,attr,omitempty"`
	LnSpc     *TextSpacing `xml:"a:lnSpc,omitempty"`
	SpcBef    *TextSpacing `xml:"a:spcBef,omitempty"`
	SpcAft    *TextSpacing `xml:"a:spcAft,omitempty"`
	BuFont    *BuFont      `xml:"a:buFont,omitempty"`
	BuNone    *BuNone      `xml:"a:buNone,omitempty"`
	BuAutoNum *BuAutoNum   `xml:"a:buAutoNum,omitempty"`
	BuChar    *BuChar      `xml:"a:buChar,omitempty"`
}

// TextSpacing is CT_TextSpacing: a spacing value expressed as either a
// percentage of single spacing (SpcPct) or an absolute size in points
// (SpcPts) — exactly one should be set. Reused for a:lnSpc, a:spcBef, and
// a:spcAft; like TextBody, it has no XMLName of its own, so the host
// struct's field tag supplies which of the three this instance is.
type TextSpacing struct {
	SpcPct *SpcPct `xml:"a:spcPct,omitempty"`
	SpcPts *SpcPts `xml:"a:spcPts,omitempty"`
}

// SpcPct is a:spcPct (CT_TextSpacingPercent): a spacing value as a
// percentage of single spacing, in thousandths of a percent (100000 = 100%).
type SpcPct struct {
	XMLName xml.Name `xml:"a:spcPct"`
	Val     int      `xml:"val,attr"`
}

// SpcPts is a:spcPts (CT_TextSpacingPoint): a spacing value as an absolute
// size, in hundredths of a point.
type SpcPts struct {
	XMLName xml.Name `xml:"a:spcPts"`
	Val     int      `xml:"val,attr"`
}

// BuFont is a:buFont (CT_TextFont): the typeface a bullet glyph is drawn
// in. Required alongside BuChar for symbol bullets — a "•" with no font
// declared renders as a missing-glyph box in some viewers.
type BuFont struct {
	XMLName  xml.Name `xml:"a:buFont"`
	Typeface string   `xml:"typeface,attr"`
}

// BuNone is a:buNone (CT_TextNoBullet): explicitly suppresses any bullet
// for this paragraph, overriding one it would otherwise inherit.
type BuNone struct {
	XMLName xml.Name `xml:"a:buNone"`
}

// BuAutoNum is a:buAutoNum (CT_TextAutonumberBullet): an automatically
// numbered bullet (e.g. "1.", "a)"), per the given ST_TextAutonumberScheme.
type BuAutoNum struct {
	XMLName xml.Name `xml:"a:buAutoNum"`
	Type    string   `xml:"type,attr"`
	StartAt int      `xml:"startAt,attr,omitempty"`
}

// BuChar is a:buChar (CT_TextCharBullet): an explicit bullet character.
type BuChar struct {
	XMLName xml.Name `xml:"a:buChar"`
	Char    string   `xml:"char,attr"`
}

// Run is a:r (CT_RegularTextRun): a single run of uniformly-formatted text.
type Run struct {
	XMLName xml.Name `xml:"a:r"`
	RPr     *RPr     `xml:"a:rPr,omitempty"`
	Text    Text     `xml:"a:t"`
}

// Text is a:t, the run's literal text. It always carries
// xml:space="preserve" so leading/trailing whitespace survives
// round-tripping through PowerPoint, mirroring docxgo's treatment of w:t.
type Text struct {
	XMLName xml.Name `xml:"a:t"`
	Space   string   `xml:"xml:space,attr"`
	Value   string   `xml:",chardata"`
}

// NewText returns a Text with xml:space="preserve" already set.
func NewText(s string) Text {
	return Text{Space: "preserve", Value: s}
}

// Br is a:br (CT_TextLineBreak): an explicit line break within a
// paragraph. PowerPoint does not treat a literal "\n" inside an a:t as a
// line break — it takes a dedicated element interleaved between runs.
type Br struct {
	XMLName xml.Name `xml:"a:br"`
}

// RPr is a:rPr (CT_TextCharacterProperties): run-level character
// formatting. Field order mirrors the schema: attributes (sz, b, i, u)
// first, then the fill group (SolidFill) ahead of the font group (Latin) —
// the validator rejects a:latin emitted before a:solidFill.
type RPr struct {
	XMLName   xml.Name   `xml:"a:rPr"`
	Sz        int        `xml:"sz,attr,omitempty"` // hundredths of a point
	B         OnOff      `xml:"b,attr,omitempty"`
	I         OnOff      `xml:"i,attr,omitempty"`
	U         string     `xml:"u,attr,omitempty"` // ST_TextUnderlineType, e.g. "sng"
	SolidFill *SolidFill `xml:"a:solidFill,omitempty"`
	Latin     *Latin     `xml:"a:latin,omitempty"`
}

// OnOff models CT_TextCharacterProperties' on/off attributes (b, i): schema
// type xsd:boolean, but Office always writes "1"/"0" rather than
// "true"/"false". It marshals only when true; the false case is handled by
// its own zero-attribute return as well as the field's omitempty tag, so
// either mechanism alone is sufficient. Explicit false ("b=\"0\"") is not
// needed yet — a future *bool would be a bigger change than Fase 2 warrants.
type OnOff bool

// MarshalXMLAttr implements xml.MarshalerAttr.
func (o OnOff) MarshalXMLAttr(name xml.Name) (xml.Attr, error) {
	if !o {
		return xml.Attr{}, nil
	}
	return xml.Attr{Name: name, Value: "1"}, nil
}

// Latin is a:latin (CT_TextFont): the Latin-script typeface.
type Latin struct {
	XMLName  xml.Name `xml:"a:latin"`
	Typeface string   `xml:"typeface,attr"`
}
