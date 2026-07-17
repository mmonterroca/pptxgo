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

	"github.com/mmonterroca/pptxgo/opc"
)

// XMLCoreProperties represents docProps/core.xml (cp:coreProperties).
// Every child is optional per the schema; only a conservative, always-safe
// subset is populated here.
type XMLCoreProperties struct {
	XMLName        xml.Name `xml:"cp:coreProperties"`
	XmlnsCP        string   `xml:"xmlns:cp,attr"`
	XmlnsDC        string   `xml:"xmlns:dc,attr"`
	XmlnsDCTerms   string   `xml:"xmlns:dcterms,attr"`
	XmlnsXSI       string   `xml:"xmlns:xsi,attr"`
	Title          string   `xml:"dc:title,omitempty"`
	Creator        string   `xml:"dc:creator,omitempty"`
	LastModifiedBy string   `xml:"cp:lastModifiedBy,omitempty"`
}

// NewCoreProperties returns docProps/core.xml content with the given title
// and creator (both optional; pass "" to omit).
func NewCoreProperties(title, creator string) *XMLCoreProperties {
	return &XMLCoreProperties{
		XmlnsCP:      opc.NamespaceCoreProperties,
		XmlnsDC:      opc.NamespaceDC,
		XmlnsDCTerms: opc.NamespaceDCTerms,
		XmlnsXSI:     opc.NamespaceXSI,
		Title:        title,
		Creator:      creator,
	}
}

// XMLAppProperties represents docProps/app.xml (Properties, extended-properties namespace).
type XMLAppProperties struct {
	XMLName     xml.Name `xml:"Properties"`
	Xmlns       string   `xml:"xmlns,attr"`
	Application string   `xml:"Application,omitempty"`
}

// NewAppProperties returns docProps/app.xml content identifying pptxgo as
// the generating application.
func NewAppProperties() *XMLAppProperties {
	return &XMLAppProperties{
		Xmlns:       opc.NamespaceExtendedProperties,
		Application: "pptxgo",
	}
}
