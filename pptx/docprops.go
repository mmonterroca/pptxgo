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
	"time"

	"github.com/mmonterroca/pptxgo/opc"
)

// Metadata holds a presentation's document properties, written to
// docProps/core.xml (the Dublin Core / OPC core properties) and
// docProps/app.xml (extended properties). Every field is optional; empty
// strings and zero times are omitted. Set it with WithMetadata, or one field
// at a time with WithTitle/WithAuthor/WithSubject/WithKeywords/
// WithDescription/WithCompany.
type Metadata struct {
	Title          string // dc:title
	Creator        string // dc:creator — the author
	Subject        string // dc:subject
	Keywords       string // cp:keywords
	Description    string // dc:description
	Category       string // cp:category
	LastModifiedBy string // cp:lastModifiedBy — defaults to Creator when empty
	Company        string // app.xml Company
	Created        time.Time
	Modified       time.Time
}

// XMLCoreProperties represents docProps/core.xml (cp:coreProperties).
// CT_CoreProperties is an xsd:all, so child order is unconstrained; every
// child is optional (omitempty). See NewCoreProperties / newCoreProperties.
type XMLCoreProperties struct {
	XMLName        xml.Name `xml:"cp:coreProperties"`
	XmlnsCP        string   `xml:"xmlns:cp,attr"`
	XmlnsDC        string   `xml:"xmlns:dc,attr"`
	XmlnsDCTerms   string   `xml:"xmlns:dcterms,attr"`
	XmlnsXSI       string   `xml:"xmlns:xsi,attr"`
	Title          string   `xml:"dc:title,omitempty"`
	Subject        string   `xml:"dc:subject,omitempty"`
	Creator        string   `xml:"dc:creator,omitempty"`
	Keywords       string   `xml:"cp:keywords,omitempty"`
	Description    string   `xml:"dc:description,omitempty"`
	LastModifiedBy string   `xml:"cp:lastModifiedBy,omitempty"`
	Category       string   `xml:"cp:category,omitempty"`
	Created        *w3cdtf  `xml:"dcterms:created,omitempty"`
	Modified       *w3cdtf  `xml:"dcterms:modified,omitempty"`
}

// w3cdtf is a dcterms date value carrying the required
// xsi:type="dcterms:W3CDTF" attribute (dcterms:created / dcterms:modified).
type w3cdtf struct {
	Type  string `xml:"xsi:type,attr"`
	Value string `xml:",chardata"`
}

// newW3CDTF renders t in the W3CDTF (ISO-8601, UTC) form OOXML dcterms dates
// use — e.g. "2026-07-19T00:00:00Z".
func newW3CDTF(t time.Time) *w3cdtf {
	return &w3cdtf{Type: "dcterms:W3CDTF", Value: t.UTC().Format("2006-01-02T15:04:05Z")}
}

// newCoreProperties builds docProps/core.xml from m. LastModifiedBy defaults
// to the creator when unset (a freshly generated file's last modifier is its
// author), matching what Office writes.
func newCoreProperties(m Metadata) *XMLCoreProperties {
	lastMod := m.LastModifiedBy
	if lastMod == "" {
		lastMod = m.Creator
	}
	cp := &XMLCoreProperties{
		XmlnsCP:        opc.NamespaceCoreProperties,
		XmlnsDC:        opc.NamespaceDC,
		XmlnsDCTerms:   opc.NamespaceDCTerms,
		XmlnsXSI:       opc.NamespaceXSI,
		Title:          m.Title,
		Subject:        m.Subject,
		Creator:        m.Creator,
		Keywords:       m.Keywords,
		Description:    m.Description,
		LastModifiedBy: lastMod,
		Category:       m.Category,
	}
	if !m.Created.IsZero() {
		cp.Created = newW3CDTF(m.Created)
	}
	if !m.Modified.IsZero() {
		cp.Modified = newW3CDTF(m.Modified)
	}
	return cp
}

// NewCoreProperties returns docProps/core.xml content with the given title
// and creator (both optional; pass "" to omit). It is a thin convenience over
// the fuller Metadata path (see WithMetadata) kept for direct callers.
func NewCoreProperties(title, creator string) *XMLCoreProperties {
	return newCoreProperties(Metadata{Title: title, Creator: creator})
}

// XMLAppProperties represents docProps/app.xml (Properties, extended-
// properties namespace). CT_Properties is a strict xsd:sequence, so field
// order matters: Company precedes Application per the schema. Both children
// are optional.
type XMLAppProperties struct {
	XMLName     xml.Name `xml:"Properties"`
	Xmlns       string   `xml:"xmlns,attr"`
	Company     string   `xml:"Company,omitempty"`
	Application string   `xml:"Application,omitempty"`
}

// newAppProperties builds docProps/app.xml, identifying pptxgo as the
// generating application and carrying the company from m.
func newAppProperties(m Metadata) *XMLAppProperties {
	return &XMLAppProperties{
		Xmlns:       opc.NamespaceExtendedProperties,
		Company:     m.Company,
		Application: "pptxgo",
	}
}

// NewAppProperties returns docProps/app.xml content identifying pptxgo as
// the generating application, with no company. Kept for direct callers; see
// WithMetadata / WithCompany for the fuller path.
func NewAppProperties() *XMLAppProperties {
	return newAppProperties(Metadata{})
}
