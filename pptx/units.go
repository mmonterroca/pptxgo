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

// NumberingScheme names an automatic bullet-numbering scheme (a:buAutoNum's
// type attribute, ST_TextAutonumberScheme) for use with Paragraph.NumberedBullet.
type NumberingScheme string

// Common numbering schemes.
const (
	NumArabicPeriod  NumberingScheme = "arabicPeriod"  // "1.", "2.", ...
	NumArabicParenR  NumberingScheme = "arabicParenR"  // "1)", "2)", ...
	NumAlphaLcPeriod NumberingScheme = "alphaLcPeriod" // "a.", "b.", ...
	NumAlphaUcPeriod NumberingScheme = "alphaUcPeriod" // "A.", "B.", ...
	NumRomanLcPeriod NumberingScheme = "romanLcPeriod" // "i.", "ii.", ...
	NumRomanUcPeriod NumberingScheme = "romanUcPeriod" // "I.", "II.", ...
)

// VerticalAnchor is a text body's vertical anchoring within its shape
// (a:bodyPr's anchor attribute), for use with ShapeRef.Anchor.
type VerticalAnchor string

// Vertical anchor positions.
const (
	AnchorTop    VerticalAnchor = "t"
	AnchorMiddle VerticalAnchor = "ctr"
	AnchorBottom VerticalAnchor = "b"
)

// AutofitMode controls how a shape's text behaves when it overflows the
// shape's bounds, for use with ShapeRef.Autofit.
type AutofitMode string

// Autofit modes.
const (
	AutofitNone        AutofitMode = "none"  // text may overflow the shape uncorrected
	AutofitShrinkText  AutofitMode = "text"  // shrink font/line-spacing to fit
	AutofitResizeShape AutofitMode = "shape" // grow the shape to fit the text
)

// PresetGeometry names a preset autoshape outline (a:prstGeom's prst
// attribute, schema type ST_ShapeType) for use with Slide.AddShape. This is
// a representative subset of the ~180 shapes ST_ShapeType allows; any other
// valid preset name can still be passed as a plain PresetGeometry("name").
type PresetGeometry string

// Common preset geometries.
const (
	ShapeRect           PresetGeometry = "rect"
	ShapeRoundRect      PresetGeometry = "roundRect"
	ShapeEllipse        PresetGeometry = "ellipse"
	ShapeTriangle       PresetGeometry = "triangle"
	ShapeRightTriangle  PresetGeometry = "rtTriangle"
	ShapeParallelogram  PresetGeometry = "parallelogram"
	ShapeTrapezoid      PresetGeometry = "trapezoid"
	ShapeDiamond        PresetGeometry = "diamond"
	ShapePentagon       PresetGeometry = "pentagon"
	ShapeHexagon        PresetGeometry = "hexagon"
	ShapeHeptagon       PresetGeometry = "heptagon"
	ShapeOctagon        PresetGeometry = "octagon"
	ShapeStar4          PresetGeometry = "star4"
	ShapeStar5          PresetGeometry = "star5"
	ShapeStar6          PresetGeometry = "star6"
	ShapeStar8          PresetGeometry = "star8"
	ShapeRightArrow     PresetGeometry = "rightArrow"
	ShapeLeftArrow      PresetGeometry = "leftArrow"
	ShapeUpArrow        PresetGeometry = "upArrow"
	ShapeDownArrow      PresetGeometry = "downArrow"
	ShapeLeftRightArrow PresetGeometry = "leftRightArrow"
	ShapeUpDownArrow    PresetGeometry = "upDownArrow"
	ShapeChevron        PresetGeometry = "chevron"
	ShapeDonut          PresetGeometry = "donut"
	ShapeNoSmoking      PresetGeometry = "noSmoking"
	ShapeHeart          PresetGeometry = "heart"
	ShapeLightningBolt  PresetGeometry = "lightningBolt"
	ShapeSun            PresetGeometry = "sun"
	ShapeMoon           PresetGeometry = "moon"
	ShapeCloud          PresetGeometry = "cloud"
	ShapeArc            PresetGeometry = "arc"
	ShapePlaque         PresetGeometry = "plaque"
	ShapeCan            PresetGeometry = "can"
	ShapeCube           PresetGeometry = "cube"
	ShapeBevel          PresetGeometry = "bevel"
	ShapeSmileyFace     PresetGeometry = "smileyFace"
	ShapeWave           PresetGeometry = "wave"
	ShapeDoubleWave     PresetGeometry = "doubleWave"
)
