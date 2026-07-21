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

// Package chart provides OpenXML modeling for DrawingML charts (c:).
package chart

import "encoding/xml"

const NamespaceMain = "http://schemas.openxmlformats.org/drawingml/2006/chart"
const NamespaceA = "http://schemas.openxmlformats.org/drawingml/2006/main"

// ChartSpace is the root element of a chart part (c:chartSpace).
type ChartSpace struct {
	XMLName xml.Name `xml:"c:chartSpace"`
	XmlnsC  string   `xml:"xmlns:c,attr"`
	XmlnsA  string   `xml:"xmlns:a,attr"`
	XmlnsR  string   `xml:"xmlns:r,attr"`

	Date1904 *Boolean `xml:"c:date1904,omitempty"`
	Chart    *Chart   `xml:"c:chart"`
}

// Chart is c:chart.
type Chart struct {
	XMLName    xml.Name    `xml:"c:chart"`
	Title      *Title      `xml:"c:title,omitempty"`
	AutoTitle  *Boolean    `xml:"c:autoTitleDeleted,omitempty"`
	PlotArea   *PlotArea   `xml:"c:plotArea"`
	Legend     *Legend     `xml:"c:legend,omitempty"`
	PlotVisOnly *Boolean   `xml:"c:plotVisOnly,omitempty"`
	DispBlanksAs *StringVal `xml:"c:dispBlanksAs,omitempty"`
}

// Title is c:title.
type Title struct {
	XMLName xml.Name `xml:"c:title"`
	Tx      *Tx      `xml:"c:tx"`
}

// Legend is c:legend.
type Legend struct {
	XMLName xml.Name `xml:"c:legend"`
	LegendPos *StringVal `xml:"c:legendPos"`
	Overlay *Boolean `xml:"c:overlay,omitempty"`
}

// PlotArea is c:plotArea.
type PlotArea struct {
	XMLName   xml.Name   `xml:"c:plotArea"`
	Layout    *Layout    `xml:"c:layout,omitempty"`
	
	// A plot area can contain one or more chart types.
	BarChart  *BarChart  `xml:"c:barChart,omitempty"`
	LineChart *LineChart `xml:"c:lineChart,omitempty"`
	PieChart  *PieChart  `xml:"c:pieChart,omitempty"`
	DoughnutChart *DoughnutChart `xml:"c:doughnutChart,omitempty"`

	// Axes
	CatAx     []*CatAx   `xml:"c:catAx,omitempty"`
	ValAx     []*ValAx   `xml:"c:valAx,omitempty"`
}

// Layout is c:layout.
type Layout struct {
	XMLName xml.Name `xml:"c:layout"`
}

// BarChart is c:barChart.
type BarChart struct {
	XMLName  xml.Name `xml:"c:barChart"`
	BarDir   *StringVal `xml:"c:barDir"` // "bar" or "col"
	Grouping *StringVal `xml:"c:grouping,omitempty"` // "clustered", "stacked", etc.
	Ser      []*Ser   `xml:"c:ser"`
	Overlap  *Overlap `xml:"c:overlap,omitempty"`
	AxId     []*UnsignedInt `xml:"c:axId"`
}

// Overlap is c:overlap.
type Overlap struct {
	XMLName xml.Name `xml:"c:overlap"`
	Val     int8     `xml:"val,attr"`
}

// LineChart is c:lineChart.
type LineChart struct {
	XMLName  xml.Name `xml:"c:lineChart"`
	Grouping *StringVal `xml:"c:grouping,omitempty"` // "standard", "stacked", etc.
	Ser      []*Ser   `xml:"c:ser"`
	AxId     []*UnsignedInt `xml:"c:axId"`
}

// PieChart is c:pieChart.
type PieChart struct {
	XMLName xml.Name `xml:"c:pieChart"`
	VaryColors *Boolean `xml:"c:varyColors,omitempty"`
	Ser      []*Ser   `xml:"c:ser"`
	// Pie charts do not use axes.
}

// DoughnutChart is c:doughnutChart.
type DoughnutChart struct {
	XMLName xml.Name `xml:"c:doughnutChart"`
	VaryColors *Boolean `xml:"c:varyColors,omitempty"`
	Ser      []*Ser   `xml:"c:ser"`
	HoleSize *UnsignedInt `xml:"c:holeSize,omitempty"`
}

// SpPr is c:spPr (Shape Properties for charts).
type SpPr struct {
	XMLName   xml.Name `xml:"c:spPr"`
	SolidFill *struct {
		XMLName xml.Name `xml:"a:solidFill"`
		SrgbClr *struct {
			XMLName xml.Name `xml:"a:srgbClr"`
			Val     string   `xml:"val,attr"`
		} `xml:"a:srgbClr,omitempty"`
		SchemeClr *struct {
			XMLName xml.Name `xml:"a:schemeClr"`
			Val     string   `xml:"val,attr"`
		} `xml:"a:schemeClr,omitempty"`
	} `xml:"a:solidFill,omitempty"`
}

// DataLabels is c:dLbls.
type DataLabels struct {
	XMLName     xml.Name `xml:"c:dLbls"`
	ShowVal     *Boolean `xml:"c:showVal,omitempty"`
	ShowCatName *Boolean `xml:"c:showCatName,omitempty"`
	ShowSerName *Boolean `xml:"c:showSerName,omitempty"`
	ShowPercent *Boolean `xml:"c:showPercent,omitempty"`
}

// Ser is c:ser (a chart series).
type Ser struct {
	XMLName xml.Name `xml:"c:ser"`
	Idx     *UnsignedInt `xml:"c:idx"`
	Order   *UnsignedInt `xml:"c:order"`
	Tx      *Tx      `xml:"c:tx,omitempty"` // Series Title
	SpPr    *SpPr    `xml:"c:spPr,omitempty"`
	DLbls   *DataLabels `xml:"c:dLbls,omitempty"`
	Cat     *Cat     `xml:"c:cat,omitempty"` // Categories (X axis)
	Val     *Val     `xml:"c:val,omitempty"` // Values (Y axis)
}

// Tx is c:tx (Text).
type Tx struct {
	XMLName xml.Name `xml:"c:tx"`
	StrRef  *StrRef  `xml:"c:strRef,omitempty"`
	V       string   `xml:"c:v,omitempty"`
}

// Cat is c:cat (Category Axis Data).
type Cat struct {
	XMLName xml.Name `xml:"c:cat"`
	StrRef  *StrRef  `xml:"c:strRef,omitempty"`
	NumRef  *NumRef  `xml:"c:numRef,omitempty"`
}

// Val is c:val (Value Axis Data).
type Val struct {
	XMLName xml.Name `xml:"c:val"`
	NumRef  *NumRef  `xml:"c:numRef"`
}

// StrRef is c:strRef (String Reference).
type StrRef struct {
	XMLName  xml.Name  `xml:"c:strRef"`
	F        string    `xml:"c:f"` // Formula
	StrCache *StrCache `xml:"c:strCache,omitempty"`
}

// StrCache is c:strCache (String Cache).
type StrCache struct {
	XMLName xml.Name `xml:"c:strCache"`
	PtCount *UnsignedInt `xml:"c:ptCount"`
	Pt      []*PtStr `xml:"c:pt"`
}

// PtStr is c:pt for strings.
type PtStr struct {
	XMLName xml.Name `xml:"c:pt"`
	Idx     uint     `xml:"idx,attr"`
	V       string   `xml:"c:v"`
}

// NumRef is c:numRef (Number Reference).
type NumRef struct {
	XMLName  xml.Name  `xml:"c:numRef"`
	F        string    `xml:"c:f"`
	NumCache *NumCache `xml:"c:numCache,omitempty"`
}

// NumCache is c:numCache (Number Cache).
type NumCache struct {
	XMLName    xml.Name `xml:"c:numCache"`
	FormatCode string   `xml:"c:formatCode"`
	PtCount    *UnsignedInt `xml:"c:ptCount"`
	Pt         []*PtNum `xml:"c:pt"`
}

// PtNum is c:pt for numbers.
type PtNum struct {
	XMLName xml.Name `xml:"c:pt"`
	Idx     uint     `xml:"idx,attr"`
	V       string   `xml:"c:v"`
}

// Boolean is a wrapper for c:val="1" or c:val="0".
type Boolean struct {
	XMLName xml.Name
	Val     uint     `xml:"val,attr"` // usually 0 or 1
}

// StringVal is a wrapper for c:val="string".
type StringVal struct {
	XMLName xml.Name
	Val     string   `xml:"val,attr"`
}

// UnsignedInt is a wrapper for c:val="uint".
type UnsignedInt struct {
	XMLName xml.Name
	Val     uint     `xml:"val,attr"`
}

// CatAx is c:catAx (Category Axis).
type CatAx struct {
	XMLName xml.Name `xml:"c:catAx"`
	AxId    *UnsignedInt `xml:"c:axId"`
	Scaling *Scaling `xml:"c:scaling"`
	Delete  *Boolean `xml:"c:delete,omitempty"`
	AxPos   *StringVal `xml:"c:axPos"`
	Title   *Title   `xml:"c:title,omitempty"`
	TickLblPos *StringVal `xml:"c:tickLblPos,omitempty"`
	CrossAx *UnsignedInt `xml:"c:crossAx"`
}

// ValAx is c:valAx (Value Axis).
type ValAx struct {
	XMLName xml.Name `xml:"c:valAx"`
	AxId    *UnsignedInt `xml:"c:axId"`
	Scaling *Scaling `xml:"c:scaling"`
	Delete  *Boolean `xml:"c:delete,omitempty"`
	AxPos   *StringVal `xml:"c:axPos"`
	MajorGridlines *struct{} `xml:"c:majorGridlines,omitempty"`
	Title   *Title   `xml:"c:title,omitempty"`
	TickLblPos *StringVal `xml:"c:tickLblPos,omitempty"`
	CrossAx *UnsignedInt `xml:"c:crossAx"`
}

// Scaling is c:scaling.
type Scaling struct {
	XMLName     xml.Name `xml:"c:scaling"`
	Orientation *StringVal `xml:"c:orientation,omitempty"` // e.g. "minMax"
}
