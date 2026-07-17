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

// Measurement conversions for DrawingML's coordinate space. EMUs (English
// Metric Units) are what Xfrm/Off/Ext and line widths are expressed in.
//
// Deliberately absent: twips (1/1440 inch), the unit WordprocessingML uses
// for paragraph spacing and indentation. PresentationML never uses twips —
// porting them here just because docxgo had them would be dead weight for
// a format that has no use for them.
const (
	// PointsPerInch is 72 points per inch.
	PointsPerInch = 72

	// EMUsPerInch is the number of English Metric Units in one inch (914400).
	EMUsPerInch = 914400

	// EMUsPerPoint is EMUsPerInch / PointsPerInch (12700). Needed for line
	// widths (a:ln w=) and any point-based positioning or sizing.
	EMUsPerPoint = 12700

	// EMUsPerCentimeter is EMUsPerInch / 2.54, rounded (360000).
	EMUsPerCentimeter = 360000
)
