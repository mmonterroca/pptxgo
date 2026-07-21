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

import (
	"encoding/xml"
	"fmt"
	"strconv"

	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/drawingml/chart"
	"github.com/mmonterroca/pptxgo/opc"
)

// ChartType defines the supported chart types.
type ChartType string

const (
	ChartTypeBar      ChartType = "bar"
	ChartTypeLine     ChartType = "line"
	ChartTypePie      ChartType = "pie"
	ChartTypeDoughnut ChartType = "doughnut"
)

// ChartBuilder is a handle onto a chart placed via Slide.AddChart.
type ChartBuilder struct {
	pres       *Presentation
	chartSpace *chart.ChartSpace
	chart      *chart.Chart
	chartType  ChartType
	series     []*chart.Ser
}

// SeriesBuilder is a handle to a chart series to configure its properties.
type SeriesBuilder struct {
	cb  *ChartBuilder
	ser *chart.Ser
}

// AddSeries adds a new data series to the chart.
func (cb *ChartBuilder) AddSeries(name string, categories []string, values []float64) *SeriesBuilder {
	idx := uint(len(cb.series))
	
	ser := &chart.Ser{
		Idx:   &chart.UnsignedInt{Val: idx},
		Order: &chart.UnsignedInt{Val: idx},
	}
	
	if name != "" {
		ser.Tx = &chart.Tx{
			StrRef: &chart.StrRef{
				F: "Sheet1!$B$1", // Dummy formula since we are read-only MVP
				StrCache: &chart.StrCache{
					PtCount: &chart.UnsignedInt{Val: 1},
					Pt: []*chart.PtStr{
						{Idx: 0, V: name},
					},
				},
			},
		}
	}
	
	if len(categories) > 0 {
		catCache := &chart.StrCache{
			PtCount: &chart.UnsignedInt{Val: uint(len(categories))},
		}
		for i, c := range categories {
			catCache.Pt = append(catCache.Pt, &chart.PtStr{Idx: uint(i), V: c})
		}
		ser.Cat = &chart.Cat{
			StrRef: &chart.StrRef{
				F: "Sheet1!$A$2:$A$" + strconv.Itoa(len(categories)+1),
				StrCache: catCache,
			},
		}
	}
	
	if len(values) > 0 {
		numCache := &chart.NumCache{
			FormatCode: "General",
			PtCount:    &chart.UnsignedInt{Val: uint(len(values))},
		}
		for i, v := range values {
			numCache.Pt = append(numCache.Pt, &chart.PtNum{
				Idx: uint(i), 
				V:   strconv.FormatFloat(v, 'f', -1, 64),
			})
		}
		ser.Val = &chart.Val{
			NumRef: &chart.NumRef{
				F: "Sheet1!$B$2:$B$" + strconv.Itoa(len(values)+1),
				NumCache: numCache,
			},
		}
	}
	
	cb.series = append(cb.series, ser)
	
	// Inject series into the correct chart element
	switch cb.chartType {
	case ChartTypeBar:
		if cb.chart.PlotArea.BarChart == nil {
			cb.chart.PlotArea.BarChart = &chart.BarChart{
				BarDir: &chart.StringVal{Val: "col"},
				Grouping: &chart.StringVal{Val: "clustered"},
				AxId: []*chart.UnsignedInt{{Val: 1}, {Val: 2}},
			}
		}
		cb.chart.PlotArea.BarChart.Ser = append(cb.chart.PlotArea.BarChart.Ser, ser)
	case ChartTypeLine:
		if cb.chart.PlotArea.LineChart == nil {
			cb.chart.PlotArea.LineChart = &chart.LineChart{
				Grouping: &chart.StringVal{Val: "standard"},
				AxId: []*chart.UnsignedInt{{Val: 1}, {Val: 2}},
			}
		}
		cb.chart.PlotArea.LineChart.Ser = append(cb.chart.PlotArea.LineChart.Ser, ser)
	case ChartTypePie:
		if cb.chart.PlotArea.PieChart == nil {
			cb.chart.PlotArea.PieChart = &chart.PieChart{
				VaryColors: &chart.Boolean{Val: 1},
			}
		}
		cb.chart.PlotArea.PieChart.Ser = append(cb.chart.PlotArea.PieChart.Ser, ser)
	case ChartTypeDoughnut:
		if cb.chart.PlotArea.DoughnutChart == nil {
			cb.chart.PlotArea.DoughnutChart = &chart.DoughnutChart{
				VaryColors: &chart.Boolean{Val: 1},
				HoleSize:   &chart.UnsignedInt{Val: 50}, // default 50%
			}
		}
		cb.chart.PlotArea.DoughnutChart.Ser = append(cb.chart.PlotArea.DoughnutChart.Ser, ser)
	}
	
	return &SeriesBuilder{cb: cb, ser: ser}
}

// Color sets a custom solid fill color (hex format, e.g., "FF0000") for the series.
func (sb *SeriesBuilder) Color(hex string) *SeriesBuilder {
	if sb.ser.SpPr == nil {
		sb.ser.SpPr = &chart.SpPr{}
	}
	sb.ser.SpPr.SolidFill = &struct {
		XMLName xml.Name `xml:"a:solidFill"`
		SrgbClr *struct {
			XMLName xml.Name `xml:"a:srgbClr"`
			Val     string   `xml:"val,attr"`
		} `xml:"a:srgbClr,omitempty"`
		SchemeClr *struct {
			XMLName xml.Name `xml:"a:schemeClr"`
			Val     string   `xml:"val,attr"`
		} `xml:"a:schemeClr,omitempty"`
	}{
		SrgbClr: &struct {
			XMLName xml.Name `xml:"a:srgbClr"`
			Val     string   `xml:"val,attr"`
		}{Val: hex},
	}
	return sb
}

// DataLabels enables data labels for this series.
func (sb *SeriesBuilder) DataLabels(showVal, showCat, showSer, showPercent bool) *SeriesBuilder {
	sb.ser.DLbls = &chart.DataLabels{}
	if showVal {
		sb.ser.DLbls.ShowVal = &chart.Boolean{Val: 1}
	} else {
		sb.ser.DLbls.ShowVal = &chart.Boolean{Val: 0}
	}
	if showCat {
		sb.ser.DLbls.ShowCatName = &chart.Boolean{Val: 1}
	}
	if showSer {
		sb.ser.DLbls.ShowSerName = &chart.Boolean{Val: 1}
	}
	if showPercent {
		sb.ser.DLbls.ShowPercent = &chart.Boolean{Val: 1}
	}
	return sb
}

// Title sets the chart's title.
func (cb *ChartBuilder) Title(title string) *ChartBuilder {
	cb.chart.Title = &chart.Title{
		Tx: &chart.Tx{
			StrRef: &chart.StrRef{
				F: "Sheet1!$B$1", // dummy
				StrCache: &chart.StrCache{
					PtCount: &chart.UnsignedInt{Val: 1},
					Pt: []*chart.PtStr{
						{Idx: 0, V: title},
					},
				},
			},
		},
	}
	return cb
}

// HasLegend enables the chart legend at the specified position (e.g., "b" for bottom, "r" for right).
func (cb *ChartBuilder) HasLegend(pos string) *ChartBuilder {
	cb.chart.Legend = &chart.Legend{
		LegendPos: &chart.StringVal{Val: pos},
		Overlay:   &chart.Boolean{Val: 0},
	}
	return cb
}

// AxisTitles sets the titles for the category (X) and value (Y) axes.
func (cb *ChartBuilder) AxisTitles(catTitle, valTitle string) *ChartBuilder {
	if cb.chartType == ChartTypePie || cb.chartType == ChartTypeDoughnut {
		return cb // Pie/Doughnut charts don't have axes
	}
	if catTitle != "" && len(cb.chartSpace.Chart.PlotArea.CatAx) > 0 {
		cb.chartSpace.Chart.PlotArea.CatAx[0].Title = &chart.Title{
			Tx: &chart.Tx{
				StrRef: &chart.StrRef{
					F: "Sheet1!$A$1",
					StrCache: &chart.StrCache{
						PtCount: &chart.UnsignedInt{Val: 1},
						Pt: []*chart.PtStr{{Idx: 0, V: catTitle}},
					},
				},
			},
		}
	}
	if valTitle != "" && len(cb.chartSpace.Chart.PlotArea.ValAx) > 0 {
		cb.chartSpace.Chart.PlotArea.ValAx[0].Title = &chart.Title{
			Tx: &chart.Tx{
				StrRef: &chart.StrRef{
					F: "Sheet1!$B$1",
					StrCache: &chart.StrCache{
						PtCount: &chart.UnsignedInt{Val: 1},
						Pt: []*chart.PtStr{{Idx: 0, V: valTitle}},
					},
				},
			},
		}
	}
	return cb
}

// SetBarDirection sets the direction of a bar chart ("col" for vertical, "bar" for horizontal).
func (cb *ChartBuilder) SetBarDirection(dir string) *ChartBuilder {
	if cb.chart.PlotArea.BarChart != nil {
		cb.chart.PlotArea.BarChart.BarDir.Val = dir
	}
	return cb
}

// SetGrouping sets the grouping of a bar/line chart ("clustered", "stacked", "percentStacked", "standard").
func (cb *ChartBuilder) SetGrouping(grouping string) *ChartBuilder {
	if cb.chart.PlotArea.BarChart != nil {
		cb.chart.PlotArea.BarChart.Grouping.Val = grouping
	}
	if cb.chart.PlotArea.LineChart != nil {
		cb.chart.PlotArea.LineChart.Grouping.Val = grouping
	}
	return cb
}

// SetHoleSize sets the hole size percentage for a doughnut chart.
func (cb *ChartBuilder) SetHoleSize(size uint) *ChartBuilder {
	if cb.chart.PlotArea.DoughnutChart != nil {
		cb.chart.PlotArea.DoughnutChart.HoleSize.Val = size
	}
	return cb
}

// AddChart adds a chart to the slide at the given position and size (x, y, w, h).
func (s *Slide) AddChart(chartType ChartType, x, y, w, h int) *ChartBuilder {
	s.pres.chartCount++
	chartID := s.pres.chartCount
	chartPartName := fmt.Sprintf("chart%d.xml", chartID)
	chartPartPath := "ppt/charts/" + chartPartName
	
	// Create chart space and plot area
	chartSpace := &chart.ChartSpace{
		XmlnsC: chart.NamespaceMain,
		XmlnsA: drawingml.NamespaceMain,
		XmlnsR: chart.NamespaceR,
		Date1904: &chart.Boolean{Val: 0},
		Chart: &chart.Chart{
			AutoTitle: &chart.Boolean{Val: 0},
			PlotVisOnly: &chart.Boolean{Val: 1},
			DispBlanksAs: &chart.StringVal{Val: "gap"},
			PlotArea: &chart.PlotArea{
				Layout: &chart.Layout{},
			},
		},
	}
	
	// Configure axes for non-pie/doughnut charts
	if chartType != ChartTypePie && chartType != ChartTypeDoughnut {
		chartSpace.Chart.PlotArea.CatAx = []*chart.CatAx{
			{
				AxId: &chart.UnsignedInt{Val: 1},
				Scaling: &chart.Scaling{
					Orientation: &chart.StringVal{Val: "minMax"},
				},
				Delete: &chart.Boolean{Val: 0},
				AxPos: &chart.StringVal{Val: "b"}, // bottom
				CrossAx: &chart.UnsignedInt{Val: 2},
				TickLblPos: &chart.StringVal{Val: "nextTo"},
			},
		}
		chartSpace.Chart.PlotArea.ValAx = []*chart.ValAx{
			{
				AxId: &chart.UnsignedInt{Val: 2},
				Scaling: &chart.Scaling{
					Orientation: &chart.StringVal{Val: "minMax"},
				},
				Delete: &chart.Boolean{Val: 0},
				AxPos: &chart.StringVal{Val: "l"}, // left
				CrossAx: &chart.UnsignedInt{Val: 1},
				MajorGridlines: &struct{}{},
			},
		}
	}
	
	// Add the chart part to the package
	s.pres.pkg.AddPart(chartPartPath, opc.ContentTypeChart, chartSpace)
	
	// Add relationship from slide to chart part
	rId, err := s.pres.pkg.Relationships(s.path).Add(opc.RelTypeChart, "../charts/"+chartPartName, "Internal")
	if err != nil {
		s.pres.errs = append(s.pres.errs, err)
	}
	
	id := s.allocID()
	
	// Create Graphic Data with c:chart
	cChart := &ChartGraphic{
		XMLName: xml.Name{Local: "c:chart"},
		XmlnsC:  chart.NamespaceMain,
		XmlnsR:  chart.NamespaceR,
		RId:     rId,
	}
	
	frame := &GraphicFrame{
		NvGraphicFramePr: &NvGraphicFramePr{
			CNvPr:             &CNvPr{ID: id, Name: fmt.Sprintf("Chart %d", id)},
			CNvGraphicFramePr: &CNvGraphicFramePr{},
			NvPr:              &NvPr{},
		},
		Xfrm: &GraphicFrameXfrm{
			Off: &drawingml.Off{X: x, Y: y},
			Ext: &drawingml.Ext{Cx: w, Cy: h},
		},
		Graphic: drawingml.NewGraphic(&drawingml.GraphicData{
			URI:   drawingml.GraphicDataURIChart,
			Inner: cChart,
		}),
	}
	
	s.spTree.Content = append(s.spTree.Content, frame)
	
	return &ChartBuilder{
		pres:       s.pres,
		chartSpace: chartSpace,
		chart:      chartSpace.Chart,
		chartType:  chartType,
	}
}

// ChartGraphic represents the c:chart element embedded within a:graphicData.
type ChartGraphic struct {
	XMLName xml.Name `xml:"c:chart"`
	XmlnsC  string   `xml:"xmlns:c,attr"`
	XmlnsR  string   `xml:"xmlns:r,attr"`
	RId     string   `xml:"r:id,attr"`
}
