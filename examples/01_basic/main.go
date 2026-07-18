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

package main

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"log"
	"os"

	"github.com/mmonterroca/pptxgo/pptx"
)

// logoPNG generates a small solid-color PNG in memory, standing in for a
// real logo/photo asset so this demo doesn't need to commit a binary file.
func logoPNG() []byte {
	img := image.NewRGBA(image.Rect(0, 0, 160, 100))
	for y := 0; y < 100; y++ {
		for x := 0; x < 160; x++ {
			img.Set(x, y, color.RGBA{R: 0x1F, G: 0x49, B: 0x7D, A: 0xFF})
		}
	}
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		log.Fatal(err)
	}
	return buf.Bytes()
}

func main() {
	p := pptx.New()

	s := p.AddSlide().BackgroundScheme(pptx.SchemeLight2)

	badge := s.AddShape(pptx.ShapeRoundRect, pptx.Inches(9.5), pptx.Inches(1), pptx.Inches(1.8), pptx.Inches(0.6)).
		FillScheme(pptx.SchemeAccent2).
		BorderScheme(pptx.SchemeDark2, 1.0)
	badge.AddParagraph().
		Text("Q3 Update").Bold().FontSize(14).Font("Calibri").ColorScheme(pptx.SchemeLight1).
		Alignment(pptx.AlignCenter)

	tb := s.AddTextBox(pptx.Inches(1), pptx.Inches(1), pptx.Inches(8), pptx.Inches(2)).
		Fill(pptx.RGB(0xE7, 0xE6, 0xE6)).
		Border(pptx.RGB(0x1F, 0x49, 0x7D), 1.5)
	tb.AddParagraph().
		Text("Quarterly Results").Bold().FontSize(32).Font("Calibri").Color(pptx.RGB(0x1F, 0x49, 0x7D)).
		Alignment(pptx.AlignCenter)

	logo := logoPNG()
	s.AddImageFromBytes(logo, pptx.Inches(1), pptx.Inches(3.5)).
		Border(pptx.RGB(0x44, 0x54, 0x6A), 1.0)
	// Same bytes as above, placed again elsewhere: pptx.Presentation dedups
	// identical media content, so this embeds only one ppt/media/ part.
	s.AddImageFromBytesWithSize(logo, pptx.Inches(11.6), pptx.Inches(6.9), pptx.Inches(0.5), pptx.Inches(0.31))

	shape := s.AddShape(pptx.ShapeEllipse, pptx.Inches(9.8), pptx.Inches(1.8), pptx.Inches(2.5), pptx.Inches(1.3)).
		Fill(pptx.RGB(0x1F, 0x49, 0x7D)).
		Border(pptx.RGB(0x44, 0x54, 0x6A), 1.0).
		BorderDash(pptx.DashDash).
		Rotation(15).
		FlipH()
	shape.AddParagraph().
		Text("On Track").Bold().FontSize(18).Font("Calibri").Color(pptx.RGB(0xFF, 0xFF, 0xFF)).
		Alignment(pptx.AlignCenter)

	list := s.AddTextBox(pptx.Inches(1), pptx.Inches(4.85), pptx.Inches(4.5), pptx.Inches(2)).
		Autofit(pptx.AutofitShrinkText).
		Insets(4, 4, 4, 4).
		Anchor(pptx.AnchorTop)
	list.AddParagraph().
		Text("Revenue up 12% year over year").FontSize(16).Font("Calibri").
		Bullet("•", "Arial").Indent(18, -18).SpaceAfter(6)
	list.AddParagraph().
		Text("Two new regions launched").FontSize(16).Font("Calibri").
		Bullet("•", "Arial").Indent(18, -18).SpaceAfter(6)
	list.AddParagraph().
		Text("Next: expand partner channel").FontSize(16).Font("Calibri").
		NumberedBullet(pptx.NumArabicPeriod).Indent(18, -18).Level(1)
	list.AddParagraph().
		Text("See the full report").FontSize(16).Font("Calibri").
		ColorScheme(pptx.SchemeHyperlink).Underline().Hyperlink("https://example.com/quarterly-report").
		Bullet("•", "Arial").Indent(18, -18)

	tbl := s.AddTable(3, 3, pptx.Inches(6), pptx.Inches(3.5), pptx.Inches(6), pptx.Inches(2.25))
	tbl.ColumnWidth(0, pptx.Inches(2.4))
	headers := []string{"Region", "Q3 Revenue", "YoY"}
	for c, h := range headers {
		tbl.Cell(0, c).Text(h).Bold()
	}
	rows := [][]string{
		{"North America", "$4.2M", "+9%"},
		{"EMEA", "$2.8M", "+15%"},
	}
	for r, row := range rows {
		for c, v := range row {
			tbl.Cell(r+1, c).Text(v)
		}
	}

	// Second slide: built from the Title and Content standard layout via
	// placeholders instead of freely-positioned shapes. Title/Body inherit
	// their geometry from slideLayout3.xml's own title/body placeholders
	// (which in turn inherit from the master's), rather than setting their
	// own a:xfrm — the inheritance chain Fase 5 exists for.
	s2 := p.AddSlide(pptx.WithLayout(pptx.LayoutTitleAndContent))
	s2.Title("Next Steps")
	body := s2.AddPlaceholder(pptx.PlaceholderBody, 1)
	body.AddParagraph().Text("Renew the partner agreement").Bullet("•", "Arial")
	body.AddParagraph().Text("Ship the Fase 5 stack").Bullet("•", "Arial")

	f, err := os.Create("01_basic_demo.pptx")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	if err := p.Save(f); err != nil {
		log.Fatal(err)
	}
}
