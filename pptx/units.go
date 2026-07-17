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

import "github.com/mmonterroca/pptxgo/drawingml"

// Inches converts f inches to EMUs. Shape positions and sizes (AddTextBox
// and friends) are all in EMUs; this and Points exist so call sites can
// name their unit instead of hand-computing the conversion.
func Inches(f float64) int {
	return int(f * drawingml.EMUsPerInch)
}

// Points converts f points to EMUs.
func Points(f float64) int {
	return int(f * drawingml.EMUsPerPoint)
}

// Emu returns n unchanged. It exists purely so a call site can spell out
// "this value is already in EMUs" instead of passing a bare int.
func Emu(n int) int {
	return n
}

// RGB constructs a drawingml.Color from 8-bit components, for use with
// Paragraph.Color.
func RGB(r, g, b uint8) drawingml.Color {
	return drawingml.Color{R: r, G: g, B: b}
}

// Alignment is a paragraph's horizontal text alignment (a:pPr's algn attribute).
type Alignment string

// Alignment values supported by Paragraph.Alignment.
const (
	AlignLeft    Alignment = "l"
	AlignCenter  Alignment = "ctr"
	AlignRight   Alignment = "r"
	AlignJustify Alignment = "just"
)
