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

// validShapeTypes is the complete ST_ShapeType enumeration: every preset
// geometry name a:prstGeom/@prst may legally hold. Extracted by reflecting
// over DocumentFormat.OpenXml.Drawing.ShapeTypeValues (OpenXML SDK 3.3.0,
// the same package PptxValidator/ already depends on) rather than
// hand-transcribed from the spec, so it's authoritative and complete — a
// hand-curated subset risks silently rejecting a legitimate shape name,
// which would be worse than the bug this validates against (AddShape
// otherwise passes any string straight into a:prstGeom/@prst; a typo
// produces a file PowerPoint refuses to open, with no error surfaced by
// Save). The pptx package's own Shape* constants (units.go) are a
// convenience subset of these same 187 values, not a separate set.
var validShapeTypes = map[PresetGeometry]bool{
	"line": true, "lineInv": true, "triangle": true, "rtTriangle": true,
	"rect": true, "diamond": true, "parallelogram": true, "trapezoid": true,
	"nonIsoscelesTrapezoid": true, "pentagon": true, "hexagon": true,
	"heptagon": true, "octagon": true, "decagon": true, "dodecagon": true,
	"star4": true, "star5": true, "star6": true, "star7": true, "star8": true,
	"star10": true, "star12": true, "star16": true, "star24": true, "star32": true,
	"roundRect": true, "round1Rect": true, "round2SameRect": true, "round2DiagRect": true,
	"snipRoundRect": true, "snip1Rect": true, "snip2SameRect": true, "snip2DiagRect": true,
	"plaque": true, "ellipse": true, "teardrop": true, "homePlate": true,
	"chevron": true, "pieWedge": true, "pie": true, "blockArc": true,
	"donut": true, "noSmoking": true, "rightArrow": true, "leftArrow": true,
	"upArrow": true, "downArrow": true, "stripedRightArrow": true, "notchedRightArrow": true,
	"bentUpArrow": true, "leftRightArrow": true, "upDownArrow": true, "leftUpArrow": true,
	"leftRightUpArrow": true, "quadArrow": true, "leftArrowCallout": true, "rightArrowCallout": true,
	"upArrowCallout": true, "downArrowCallout": true, "leftRightArrowCallout": true,
	"upDownArrowCallout": true, "quadArrowCallout": true, "bentArrow": true,
	"uturnArrow": true, "circularArrow": true, "leftCircularArrow": true,
	"leftRightCircularArrow": true, "curvedRightArrow": true, "curvedLeftArrow": true,
	"curvedUpArrow": true, "curvedDownArrow": true, "swooshArrow": true,
	"cube": true, "can": true, "lightningBolt": true, "heart": true,
	"sun": true, "moon": true, "smileyFace": true, "irregularSeal1": true,
	"irregularSeal2": true, "foldedCorner": true, "bevel": true, "frame": true,
	"halfFrame": true, "corner": true, "diagStripe": true, "chord": true,
	"arc": true, "leftBracket": true, "rightBracket": true, "leftBrace": true,
	"rightBrace": true, "bracketPair": true, "bracePair": true,
	"straightConnector1": true, "bentConnector2": true, "bentConnector3": true,
	"bentConnector4": true, "bentConnector5": true, "curvedConnector2": true,
	"curvedConnector3": true, "curvedConnector4": true, "curvedConnector5": true,
	"callout1": true, "callout2": true, "callout3": true, "accentCallout1": true,
	"accentCallout2": true, "accentCallout3": true, "borderCallout1": true,
	"borderCallout2": true, "borderCallout3": true, "accentBorderCallout1": true,
	"accentBorderCallout2": true, "accentBorderCallout3": true,
	"wedgeRectCallout": true, "wedgeRoundRectCallout": true, "wedgeEllipseCallout": true,
	"cloudCallout": true, "cloud": true, "ribbon": true, "ribbon2": true,
	"ellipseRibbon": true, "ellipseRibbon2": true, "leftRightRibbon": true,
	"verticalScroll": true, "horizontalScroll": true, "wave": true, "doubleWave": true,
	"plus": true, "flowChartProcess": true, "flowChartDecision": true,
	"flowChartInputOutput": true, "flowChartPredefinedProcess": true,
	"flowChartInternalStorage": true, "flowChartDocument": true,
	"flowChartMultidocument": true, "flowChartTerminator": true,
	"flowChartPreparation": true, "flowChartManualInput": true,
	"flowChartManualOperation": true, "flowChartConnector": true,
	"flowChartPunchedCard": true, "flowChartPunchedTape": true,
	"flowChartSummingJunction": true, "flowChartOr": true, "flowChartCollate": true,
	"flowChartSort": true, "flowChartExtract": true, "flowChartMerge": true,
	"flowChartOfflineStorage": true, "flowChartOnlineStorage": true,
	"flowChartMagneticTape": true, "flowChartMagneticDisk": true,
	"flowChartMagneticDrum": true, "flowChartDisplay": true, "flowChartDelay": true,
	"flowChartAlternateProcess": true, "flowChartOffpageConnector": true,
	"actionButtonBlank": true, "actionButtonHome": true, "actionButtonHelp": true,
	"actionButtonInformation": true, "actionButtonForwardNext": true,
	"actionButtonBackPrevious": true, "actionButtonEnd": true, "actionButtonBeginning": true,
	"actionButtonReturn": true, "actionButtonDocument": true, "actionButtonSound": true,
	"actionButtonMovie": true, "gear6": true, "gear9": true, "funnel": true,
	"mathPlus": true, "mathMinus": true, "mathMultiply": true, "mathDivide": true,
	"mathEqual": true, "mathNotEqual": true, "cornerTabs": true, "squareTabs": true,
	"plaqueTabs": true, "chartX": true, "chartStar": true, "chartPlus": true,
}

// IsValidPresetGeometry reports whether prst is one of ST_ShapeType's 187
// defined preset geometry names.
func IsValidPresetGeometry(prst PresetGeometry) bool {
	return validShapeTypes[prst]
}
