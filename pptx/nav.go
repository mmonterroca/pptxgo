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
	"strconv"
)

// xmlPresentationNav is a READ-SAFE mirror of just enough of p:presentation
// (CT_Presentation) to resolve slide order — deliberately separate from
// XMLPresentation (xml.go), which is write-only: a literal "p:presentation"
// tag (as XMLPresentation uses) never matches on unmarshal, since Go's
// encoding/xml treats the text before the colon as part of the local name,
// not a namespace prefix, so it never resolves to the actual namespace URI
// a real decoder assigns the "p:" prefix to. Local-name-only tags here
// match regardless of which prefix the source document happens to declare
// for the presentationml namespace.
type xmlPresentationNav struct {
	XMLName  xml.Name        `xml:"presentation"`
	SldIdLst *xmlSldIdLstNav `xml:"sldIdLst"`
}

// xmlSldIdLstNav is p:sldIdLst's read-safe mirror: the ordered list of
// slides in the presentation.
type xmlSldIdLstNav struct {
	Entries []xmlSldIdNav `xml:"sldId"`
}

// officeDocumentRelsNS is the officeDocument relationships namespace, as a
// literal rather than a reference to opc.NamespaceOfficeDocumentRels
// (which it must stay equal to — see the test asserting that): it is
// compared against an xml.Name.Space at runtime in UnmarshalXML below, a
// context an actual Go constant would serve just as well as a literal, but
// keeping it a literal here mirrors why the equivalent struct-tag approach
// (tried first — see xmlSldIdNav's own doc comment) needed one.
const officeDocumentRelsNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

// xmlSldIdNav is one p:sldId entry's read-safe mirror. It implements
// UnmarshalXML itself rather than relying on struct-tag attribute matching
// because a real sldId element carries two attributes sharing the local
// name "id" in different namespaces — the bare "id" (ST_SlideId, no
// namespace) and "r:id" (the relationship reference, in the officeDocument
// relationships namespace) — and Go's encoding/xml resolves an
// UNQUALIFIED attr tag (e.g. "id,attr") by local name ALONE, ignoring
// namespace entirely; there is no struct-tag syntax for "this local name,
// but only with no namespace". With both "id" and "r:id" satisfying an
// unqualified "id,attr" tag, whichever attribute appears LAST in the
// source document silently overwrites the field — confirmed empirically
// (a real sldId's r:id always follows id in document order, so ID was
// always getting the r:id STRING value, failing ParseUint). Explicit
// per-attribute namespace comparison in UnmarshalXML is the only reliable
// fix.
type xmlSldIdNav struct {
	ID  uint64
	RID string
}

// UnmarshalXML implements xml.Unmarshaler. sldId has no child elements, so
// this only needs to walk start's own attributes and skip the (empty)
// element body.
func (s *xmlSldIdNav) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch {
		case attr.Name.Local == "id" && attr.Name.Space == "":
			id, err := strconv.ParseUint(attr.Value, 10, 64)
			if err != nil {
				return err
			}
			s.ID = id
		case attr.Name.Local == "id" && attr.Name.Space == officeDocumentRelsNS:
			s.RID = attr.Value
		}
	}
	return d.Skip()
}
