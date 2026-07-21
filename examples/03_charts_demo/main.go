package main

import (
	"log"
	"os"

	"github.com/mmonterroca/pptxgo/pptx"
)

func main() {
	pres := pptx.New()

	slide1 := pres.AddSlide()
	slide1.AddTextBox(pptx.Inches(1), pptx.Inches(0.5), pptx.Inches(8), pptx.Inches(1)).
		AddParagraph().Text("Bar Chart Demo").
		FontSize(24).Bold()

	// Add a Stacked Horizontal Bar Chart with Axis Titles
	chart1 := slide1.AddChart(pptx.ChartTypeBar, pptx.Inches(1), pptx.Inches(1.5), pptx.Inches(8), pptx.Inches(4))
	chart1.Title("Quarterly Performance").HasLegend("b").AxisTitles("Categories", "Millions USD")
	chart1.SetBarDirection("bar").SetGrouping("stacked")
	chart1.AddSeries("Series 1", []string{"Category 1", "Category 2", "Category 3"}, []float64{4.3, 2.5, 3.5}).Color("00FF00").DataLabels(true, false, false, false)
	chart1.AddSeries("Series 2", []string{"Category 1", "Category 2", "Category 3"}, []float64{2.4, 4.4, 1.8}).Color("0000FF").DataLabels(true, false, false, false)

	slide2 := pres.AddSlide()
	slide2.AddTextBox(pptx.Inches(1), pptx.Inches(0.5), pptx.Inches(8), pptx.Inches(1)).
		AddParagraph().Text("Pie Chart Demo").
		FontSize(24).Bold()

	// Add a Pie Chart
	chart2 := slide2.AddChart(pptx.ChartTypePie, pptx.Inches(1), pptx.Inches(1.5), pptx.Inches(4), pptx.Inches(4))
	chart2.Title("Sales by Region").HasLegend("b")
	chart2.AddSeries("Sales", []string{"Q1", "Q2", "Q3", "Q4"}, []float64{8.2, 3.2, 1.4, 1.2}).DataLabels(true, false, false, true)

	// Add a Doughnut Chart next to it
	chart3 := slide2.AddChart(pptx.ChartTypeDoughnut, pptx.Inches(5.5), pptx.Inches(1.5), pptx.Inches(4), pptx.Inches(4))
	chart3.Title("Market Share").HasLegend("b").SetHoleSize(60)
	chart3.AddSeries("Share", []string{"Alpha", "Beta", "Gamma"}, []float64{45.5, 30.2, 24.3}).DataLabels(true, false, false, false)

	f, err := os.Create("03_charts_demo.pptx")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	if err := pres.Save(f); err != nil {
		log.Fatalf("Save failed: %v", err)
	}
	log.Println("Created 03_charts_demo.pptx successfully!")
}
