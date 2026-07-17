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

// NamespaceMain is the DrawingML main namespace URI ("a:").
const NamespaceMain = "http://schemas.openxmlformats.org/drawingml/2006/main"

// GraphicDataURITable is the graphicData uri identifying table content
// (an a:tbl inside GraphicData.Inner).
const GraphicDataURITable = "http://schemas.openxmlformats.org/drawingml/2006/table"

// GraphicDataURIChart is the graphicData uri identifying chart content
// (a c:chart inside GraphicData.Inner).
const GraphicDataURIChart = "http://schemas.openxmlformats.org/drawingml/2006/chart"

// Graphic is a:graphic: the outer container for graphic-frame content
// (tables, charts, embedded objects, SmartArt) in both WordprocessingML
// and PresentationML. It is not used to embed a plain picture in
// PresentationML — see the package doc comment.
type Graphic struct {
	XMLName     xml.Name     `xml:"a:graphic"`
	Xmlns       string       `xml:"xmlns:a,attr"`
	GraphicData *GraphicData `xml:"a:graphicData"`
}

// NewGraphic wraps data in a Graphic with the standard DrawingML namespace
// declaration.
func NewGraphic(data *GraphicData) *Graphic {
	return &Graphic{Xmlns: NamespaceMain, GraphicData: data}
}

// GraphicData is a:graphicData: identifies, via URI, what kind of content
// this graphic frame holds. Inner holds that content (an a:tbl, a c:chart,
// ...) and is left as `any` — drawingml does not model tables or charts;
// those belong to the packages that do.
type GraphicData struct {
	XMLName xml.Name `xml:"a:graphicData"`
	URI     string   `xml:"uri,attr"`
	Inner   any      `xml:",any"`
}
