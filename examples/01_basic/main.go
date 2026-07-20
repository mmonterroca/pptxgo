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
	"time"

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

// brandTheme starts from the default Office theme and overrides just the
// brand-relevant slots (accent palette + heading/body fonts). Because every
// shape below references colors by scheme slot (FillScheme, BorderScheme,
// ColorScheme, BackgroundScheme), this one Theme recolors all of them at
// once — the badge's SchemeAccent2, the ellipse-adjacent SchemeDark2 border,
// the SchemeLight2 background, and the SchemeHyperlink link color.
func brandTheme() pptx.Theme {
	t := pptx.DefaultTheme()
	t.Name = "pptxgo Demo"
	t.Colors.Dark2 = pptx.RGB(0x1F, 0x49, 0x7D)   // deep navy
	t.Colors.Accent1 = pptx.RGB(0x1F, 0x49, 0x7D) // deep navy
	t.Colors.Accent2 = pptx.RGB(0xED, 0x7D, 0x31) // warm orange badge
	t.Colors.Hyperlink = pptx.RGB(0x1F, 0x49, 0x7D)
	t.Fonts.Major = "Calibri Light"
	t.Fonts.Minor = "Calibri"
	return t
}

func main() {
	// A fixed timestamp keeps the generated file deterministic (so the
	// validated demo output is reproducible); real callers would use
	// time.Now().
	authored := time.Date(2026, 7, 19, 0, 0, 0, 0, time.UTC)
	p := pptx.New(
		pptx.WithTheme(brandTheme()),
		pptx.WithMetadata(pptx.Metadata{
			Title:    "Quarterly Results",
			Creator:  "pptxgo demo",
			Subject:  "Q3 business review",
			Keywords: "quarterly, revenue, demo",
			Company:  "pptxgo",
			Created:  authored,
			Modified: authored,
		}),
	)

	s := p.AddSlide().BackgroundScheme(pptx.SchemeLight2)

	badge := s.AddShape(pptx.ShapeRoundRect, pptx.Inches(9.5), pptx.Inches(1), pptx.Inches(1.8), pptx.Inches(0.6)).
		FillScheme(pptx.SchemeAccent2).
		BorderScheme(pptx.SchemeDark2, 1.0).
		// Adjust the roundRect's corner-radius handle to 30% of the short side.
		Adjust("adj", 30000)
	badge.AddParagraph().
		Text("Q3 Update").Bold().FontSize(14).Font("Calibri").ColorScheme(pptx.SchemeLight1).
		Alignment(pptx.AlignCenter)

	// Gradient stops referenced by theme slot (not hardcoded RGB), so this
	// badge recolors along with the rest of the deck when the theme changes.
	// Shade/Tint on the stops themselves darken/lighten each end of the
	// blend without hardcoding a second literal color.
	trending := s.AddShape(pptx.ShapeRoundRect, pptx.Inches(11.5), pptx.Inches(1), pptx.Inches(1.6), pptx.Inches(0.6)).
		Adjust("adj", 30000).
		GradientFill(45,
			pptx.GradientStop{Scheme: pptx.SchemeAccent2, Pos: 0, Shade: 20},
			pptx.GradientStop{Scheme: pptx.SchemeAccent4, Pos: 100, Tint: 30}).
		Glow(pptx.RGB(0xED, 0x7D, 0x31), 4)
	trending.AddParagraph().
		Text("Trending Up").Bold().FontSize(14).Font("Calibri").Color(pptx.RGB(0xFF, 0xFF, 0xFF)).
		Alignment(pptx.AlignCenter)

	// LineCap/LineJoin round the rectangle's own corners and line ends;
	// Shadow adds Office's own default drop-shadow preset behind it.
	tb := s.AddTextBox(pptx.Inches(1), pptx.Inches(1), pptx.Inches(8), pptx.Inches(2)).
		Fill(pptx.RGB(0xE7, 0xE6, 0xE6)).
		Border(pptx.RGB(0x1F, 0x49, 0x7D), 1.5).
		LineCap(pptx.LineCapRound).
		LineJoin(pptx.LineJoinRound).
		Shadow(pptx.RGB(0x1F, 0x49, 0x7D), 40)
	tb.AddParagraph().
		Text("Quarterly Results").Bold().FontSize(32).Font("Calibri").Color(pptx.RGB(0x1F, 0x49, 0x7D)).
		Alignment(pptx.AlignCenter)

	// A thin vertical divider between the left-hand content and the table,
	// modeled as an open "line" preset (not a p:cxnSp connector — those are
	// out of scope) so ArrowEnd/LineCap have visible effect: a closed
	// autoshape's outline has no defined start/end for an arrowhead to sit on.
	s.AddShape(pptx.ShapeLine, pptx.Inches(5.7), pptx.Inches(1), pptx.Emu(1), pptx.Inches(5.9)).
		Border(pptx.RGB(0x44, 0x54, 0x6A), 1.5).
		LineCap(pptx.LineCapRound).
		ArrowEnd(pptx.ArrowheadTriangle)

	logo := logoPNG()
	// SoftEdges fades the image's own edges to transparent — shown WITHOUT a
	// Border here, since a hard outline would just redraw a crisp edge over
	// the fade and mask the effect.
	s.AddImageFromBytes(logo, pptx.Inches(1), pptx.Inches(3.5)).
		SoftEdges(8)
	// Same bytes as above, placed again elsewhere: pptx.Presentation dedups
	// identical media content, so this embeds only one ppt/media/ part.
	s.AddImageFromBytesWithSize(logo, pptx.Inches(11.6), pptx.Inches(6.9), pptx.Inches(0.5), pptx.Inches(0.31))

	shape := s.AddShape(pptx.ShapeEllipse, pptx.Inches(9.8), pptx.Inches(1.8), pptx.Inches(2.5), pptx.Inches(1.3)).
		Fill(pptx.RGB(0x1F, 0x49, 0x7D)).
		Border(pptx.RGB(0x44, 0x54, 0x6A), 1.0).
		BorderDash(pptx.DashDash).
		Rotation(15).
		FlipH().
		Reflection(35)
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
		NumberedBulletFrom(pptx.NumArabicPeriod, 2).Indent(18, -18).Level(1)
	list.AddParagraph().
		Text("See the full report").FontSize(16).Font("Calibri").
		ColorScheme(pptx.SchemeHyperlink).Underline().Hyperlink("https://example.com/quarterly-report").
		Bullet("•", "Arial").Indent(18, -18)

	tbl := s.AddTable(4, 3, pptx.Inches(6), pptx.Inches(3.5), pptx.Inches(6), pptx.Inches(2.9))
	tbl.ColumnWidth(0, pptx.Inches(2.4))
	headers := []string{"Region", "Q3 Revenue", "YoY"}
	for c, h := range headers {
		cell := tbl.Cell(0, c)
		// Branded header row: theme-slot cell fill + light text + centered
		// vertically, all via TableCell's TcPr — the fills follow WithTheme.
		// BorderScheme underlines the row with a themed per-side border.
		cell.FillScheme(pptx.SchemeAccent1).Anchor(pptx.AnchorMiddle).
			BorderScheme(pptx.SideBottom, pptx.SchemeDark2, 1.0)
		cell.Text(h).Bold().ColorScheme(pptx.SchemeLight1)
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
	// Merge the "Total" row's first two columns into a single labeled cell —
	// pptx.Table.MergeCells (gridSpan/hMerge on the surviving cells, no
	// <a:tc> ever deleted).
	tbl.MergeCells(3, 0, 3, 1)
	tbl.Cell(3, 0).Border(pptx.SideTop, pptx.RGB(0x1F, 0x49, 0x7D), 1.5).Text("Total").Bold()
	tbl.Cell(3, 2).Border(pptx.SideTop, pptx.RGB(0x1F, 0x49, 0x7D), 1.5).Text("+11%").Bold()

	// Speaker notes for this slide — shown in PowerPoint's notes pane and on
	// the printed notes page. The first Notes call lazily creates the deck's
	// single notes master.
	s.Notes("Lead with the +11% total.\nEMEA is the standout at +15% YoY.")

	// Second slide: built from the Title and Content standard layout via
	// placeholders instead of freely-positioned shapes. Title/Body inherit
	// their geometry from slideLayout3.xml's own title/body placeholders
	// (which in turn inherit from the master's), rather than setting their
	// own a:xfrm — the inheritance chain Fase 5 exists for.
	s2 := p.AddSlide(pptx.WithLayout(pptx.LayoutTitleAndContent)).
		BackgroundGradient(90,
			pptx.GradientStop{Color: pptx.RGB(0xFF, 0xFF, 0xFF), Pos: 0},
			pptx.GradientStop{Color: pptx.RGB(0xDC, 0xE6, 0xF1), Pos: 100})
	s2.Title("Next Steps")
	body := s2.AddPlaceholder(pptx.PlaceholderBody, 1)
	body.AddParagraph().Text("Renew the partner agreement").Bullet("•", "Arial")
	body.AddParagraph().Text("Ship the Fase 5 stack").Bullet("•", "Arial")
	// No explicit Bullet/Indent below — each paragraph inherits its bullet
	// glyph and indent entirely from the master's own multi-level txStyles
	// cascade (NewDefaultTxStyles' 9-level bodyStyle), proving the
	// inheritance actually resolves per level, not just at level 0.
	body.AddParagraph().Text("Expand partner channel").Level(0)
	body.AddParagraph().Text("Identify regional partners").Level(1)
	body.AddParagraph().Text("Confirm SLAs with each partner").Level(2)

	s2.Notes("Close by assigning owners to each next step before the meeting ends.")

	// Footer row: a literal date (bottom-left), footer text (bottom-center),
	// and an auto-updating slide-number field (bottom-right). Each is a
	// self-positioned placeholder, so it renders without a master hf entry.
	s2.DateText("July 19, 2026").Footer("Acme — Confidential").SlideNumber()

	// Third slide: a native diagram — a group of three shapes that move,
	// resize, and rotate together in PowerPoint's own UI (Slide.AddGroup),
	// linked by connectors bound to each shape's own connection sites
	// (Slide.Connect) — unlike a plain AddShape(ShapeLine, ...) + arrowhead
	// (the divider on slide 1), a bound connector follows its shapes when
	// they move in PowerPoint. LibreOffice is NOT authoritative for a
	// connector's binding, only its routing looks visually right there —
	// open in real PowerPoint and drag a box to confirm the connector
	// actually follows.
	s3 := p.AddSlide()
	s3.AddTextBox(pptx.Inches(1), pptx.Inches(0.5), pptx.Inches(8), pptx.Inches(0.8)).
		AddParagraph().Text("Process Flow").Bold().FontSize(28).Font("Calibri").Color(pptx.RGB(0x1F, 0x49, 0x7D))

	// The group's own bounding box roughly encloses its three members —
	// AddGroup doesn't compute this automatically (see its own doc
	// comment), so it's picked to fit the layout below.
	flow := s3.AddGroup(pptx.Inches(0.8), pptx.Inches(1.8), pptx.Inches(10.4), pptx.Inches(1.4))
	discovery := flow.AddShape(pptx.ShapeRoundRect, pptx.Inches(1), pptx.Inches(2), pptx.Inches(2), pptx.Inches(1)).
		FillScheme(pptx.SchemeAccent1)
	discovery.AddParagraph().Text("Discovery").Bold().FontSize(14).Font("Calibri").ColorScheme(pptx.SchemeLight1).Alignment(pptx.AlignCenter)
	design := flow.AddShape(pptx.ShapeRoundRect, pptx.Inches(5), pptx.Inches(2), pptx.Inches(2), pptx.Inches(1)).
		FillScheme(pptx.SchemeAccent1)
	design.AddParagraph().Text("Design").Bold().FontSize(14).Font("Calibri").ColorScheme(pptx.SchemeLight1).Alignment(pptx.AlignCenter)
	ship := flow.AddShape(pptx.ShapeRoundRect, pptx.Inches(9), pptx.Inches(2), pptx.Inches(2), pptx.Inches(1)).
		FillScheme(pptx.SchemeAccent1)
	ship.AddParagraph().Text("Ship").Bold().FontSize(14).Font("Calibri").ColorScheme(pptx.SchemeLight1).Alignment(pptx.AlignCenter)

	// Bound connectors between group members — Discovery's own right-side
	// connection site to Design's own left, and Design's right to Ship's
	// left.
	s3.Connect(discovery, pptx.SiteRight, design, pptx.SiteLeft, pptx.ConnStraight).
		BorderScheme(pptx.SchemeDark2, 1.5).
		ArrowEnd(pptx.ArrowheadTriangle)
	s3.Connect(design, pptx.SiteRight, ship, pptx.SiteLeft, pptx.ConnStraight).
		BorderScheme(pptx.SchemeDark2, 1.5).
		ArrowEnd(pptx.ArrowheadTriangle)

	// A shape OUTSIDE the group, connected to a shape INSIDE it — proves
	// connector binding works by shape id (slide-global) regardless of
	// group nesting, not just between top-level or same-group siblings.
	gate := s3.AddShape(pptx.ShapeEllipse, pptx.Inches(4.5), pptx.Inches(4.2), pptx.Inches(3), pptx.Inches(1)).
		FillScheme(pptx.SchemeAccent2).
		BorderScheme(pptx.SchemeDark2, 1.0)
	gate.AddParagraph().Text("Review Gate").Bold().FontSize(14).Font("Calibri").ColorScheme(pptx.SchemeLight1).Alignment(pptx.AlignCenter)
	s3.Connect(design, pptx.SiteBottom, gate, pptx.SiteTop, pptx.ConnBent).
		BorderScheme(pptx.SchemeDark2, 1.5).
		LineJoin(pptx.LineJoinRound).
		ArrowEnd(pptx.ArrowheadTriangle)

	f, err := os.Create("01_basic_demo.pptx")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	if err := p.Save(f); err != nil {
		log.Fatal(err)
	}
}
