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

// StCxn is a:stCxn (CT_Connection): binds a connector's START point to a
// connection site on another shape — ID is that shape's own p:cNvPr/@id,
// Idx is the connection-site index on its geometry (ST_ShapeType's built-in
// autoshapes number these 0=top-center, 1=left-center, 2=bottom-center,
// 3=right-center, counter-clockwise from the top — confirmed against
// python-pptx's own hardcoded convention and a real PowerPoint-compatible
// render, not derived from the schema alone, the same "extract from a real
// file" discipline Table.MergeCells' encoding already follows). EndCxn is
// the same shape for the connector's END point. A separate type from
// EndCxn, not a reuse of it, even though both are schema-identical
// CT_Connection: a fixed XMLName can't be renamed via a field tag — the
// same conflict ChOff/TcBorderLn/LineEnd/GraphicFrameXfrm's own doc
// comments already document.
type StCxn struct {
	XMLName xml.Name `xml:"a:stCxn"`
	ID      uint32   `xml:"id,attr"`
	Idx     uint32   `xml:"idx,attr"`
}

// EndCxn is a:endCxn (CT_Connection) — see StCxn's own doc comment for the
// connection-site index convention and why this can't reuse StCxn directly.
type EndCxn struct {
	XMLName xml.Name `xml:"a:endCxn"`
	ID      uint32   `xml:"id,attr"`
	Idx     uint32   `xml:"idx,attr"`
}
