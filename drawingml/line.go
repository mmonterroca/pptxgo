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

// Ln is a:ln (CT_LineProperties): a shape or picture's outline. Only width
// and a solid-color fill are modeled — the schema's other line properties
// (dash pattern, cap, join, head/tail decorations) are out of scope until a
// caller needs them.
type Ln struct {
	XMLName   xml.Name   `xml:"a:ln"`
	W         int        `xml:"w,attr,omitempty"` // EMUs
	SolidFill *SolidFill `xml:"a:solidFill,omitempty"`
}

// NewLn returns a solid-color outline of the given width in EMUs.
func NewLn(c Color, widthEMU int) *Ln {
	return &Ln{W: widthEMU, SolidFill: NewSolidFillRGB(c)}
}

// NewLnScheme returns a theme-color outline of the given width in EMUs.
func NewLnScheme(schemeColor string, widthEMU int) *Ln {
	return &Ln{W: widthEMU, SolidFill: NewSolidFillScheme(schemeColor)}
}
