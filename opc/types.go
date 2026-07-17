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

package opc

import "encoding/xml"

// XMLRelationships represents the root element of a .rels part.
type XMLRelationships struct {
	XMLName       xml.Name           `xml:"Relationships"`
	Xmlns         string             `xml:"xmlns,attr"`
	Relationships []*XMLRelationship `xml:"Relationship"`
}

// XMLRelationship represents a single Relationship element within a .rels part.
type XMLRelationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

// XMLContentTypes represents [Content_Types].xml.
type XMLContentTypes struct {
	XMLName   xml.Name       `xml:"Types"`
	Xmlns     string         `xml:"xmlns,attr"`
	Defaults  []*XMLDefault  `xml:"Default"`
	Overrides []*XMLOverride `xml:"Override"`
}

// XMLDefault maps a file extension to a content type.
type XMLDefault struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// XMLOverride maps a specific part path to a content type, overriding any
// Default that would otherwise apply by extension.
type XMLOverride struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}
