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

// Blip (a:blip) references an embedded or linked image by relationship ID.
// It carries its own xmlns:r declaration so it is self-contained wherever
// it is embedded, rather than depending on an ancestor element to have
// declared the "r" prefix.
type Blip struct {
	XMLName xml.Name `xml:"a:blip"`
	XmlnsR  string   `xml:"xmlns:r,attr"`
	Embed   string   `xml:"r:embed,attr"`
}

// NamespaceRelationships is the officeDocument relationships namespace,
// used for the "r:" prefix on Blip.Embed.
const NamespaceRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

// NewBlip creates a Blip referencing the given relationship ID.
func NewBlip(relID string) *Blip {
	return &Blip{XmlnsR: NamespaceRelationships, Embed: relID}
}

// Stretch (a:stretch) is the simplest image fill mode: stretch the image to
// fill its containing rectangle.
type Stretch struct {
	XMLName  xml.Name  `xml:"a:stretch"`
	FillRect *FillRect `xml:"a:fillRect,omitempty"`
}

// FillRect (a:fillRect) is the rectangle an image fill stretches to. An
// empty element means "the whole shape".
type FillRect struct {
	XMLName xml.Name `xml:"a:fillRect"`
}

// PicLocks (a:picLocks) restricts what a user may do to a picture in the
// authoring UI.
type PicLocks struct {
	XMLName            xml.Name `xml:"a:picLocks"`
	NoChangeAspect     bool     `xml:"noChangeAspect,attr,omitempty"`
	NoChangeArrowheads bool     `xml:"noChangeArrowheads,attr,omitempty"`
}
