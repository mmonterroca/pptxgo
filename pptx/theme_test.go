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
	"strings"
	"testing"

	"github.com/mmonterroca/pptxgo/drawingml"
)

// clrSchemeReadback mirrors a:clrScheme for reading the rendered theme back
// out and asserting every slot resolved to a color — dk1/lt1 as sysClr, the
// rest as srgbClr.
type clrSchemeReadback struct {
	XMLName xml.Name `xml:"clrScheme"`
	Name    string   `xml:"name,attr"`
	Slots   []struct {
		XMLName xml.Name
		Srgb    struct {
			Val string `xml:"val,attr"`
		} `xml:"srgbClr"`
		Sys struct {
			Val string `xml:"val,attr"`
		} `xml:"sysClr"`
	} `xml:",any"`
}

func TestDefaultTheme_RendersOfficePaletteAndFonts(t *testing.T) {
	xmlBytes := renderThemeXML(DefaultTheme())

	// Well-formed.
	var probe any
	if err := xml.Unmarshal(xmlBytes, &probe); err != nil {
		t.Fatalf("rendered default theme is not well-formed XML: %v", err)
	}

	s := string(xmlBytes)
	for _, want := range []string{
		`<a:sysClr val="windowText" lastClr="000000">`, // dk1 (system-adaptive, Office convention)
		`<a:sysClr val="window" lastClr="FFFFFF">`,     // lt1
		`<a:srgbClr val="44546A">`,                     // dk2
		`<a:srgbClr val="4472C4">`,                     // accent1
		`<a:srgbClr val="ED7D31">`,                     // accent2
		`<a:srgbClr val="0563C1">`,                     // hlink
		`<a:srgbClr val="954F72">`,                     // folHlink
		`typeface="Calibri Light"`,                     // major font
		`typeface="Calibri"`,                           // minor font
		`<a:fmtScheme name="Office">`,
	} {
		if !strings.Contains(s, want) {
			t.Errorf("expected default theme to contain %q, got:\n%s", want, s)
		}
	}
}

func TestRenderThemeXML_HasAllTwelveColorSlotsInSchemaOrder(t *testing.T) {
	var scheme clrSchemeReadback
	// Extract just the a:clrScheme fragment to unmarshal (namespace-prefixed
	// local names resolve by local name here).
	full := string(renderThemeXML(DefaultTheme()))
	start := strings.Index(full, "<a:clrScheme")
	end := strings.Index(full, "</a:clrScheme>") + len("</a:clrScheme>")
	if start == -1 || end < len("</a:clrScheme>") {
		t.Fatalf("could not locate a:clrScheme fragment in %s", full)
	}
	if err := xml.Unmarshal([]byte(full[start:end]), &scheme); err != nil {
		t.Fatalf("unmarshal clrScheme: %v", err)
	}

	wantOrder := []string{"dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"}
	if len(scheme.Slots) != len(wantOrder) {
		t.Fatalf("expected %d color slots, got %d", len(wantOrder), len(scheme.Slots))
	}
	for i, want := range wantOrder {
		if scheme.Slots[i].XMLName.Local != want {
			t.Errorf("slot %d: expected element %q, got %q", i, want, scheme.Slots[i].XMLName.Local)
		}
		// dk1/lt1 are sysClr; every other slot is srgbClr. Each must carry
		// exactly one non-empty color.
		if scheme.Slots[i].Srgb.Val == "" && scheme.Slots[i].Sys.Val == "" {
			t.Errorf("slot %d (%s): expected a non-empty srgbClr or sysClr val", i, want)
		}
	}
}

func TestWithTheme_SwapsColorsAndFontsInThemePart(t *testing.T) {
	brand := DefaultTheme()
	brand.Name = "Acme"
	brand.Colors.Accent1 = drawingml.Color{R: 0x1F, G: 0x49, B: 0x7D}
	brand.Fonts.Minor = "Montserrat"
	brand.Fonts.Major = "Montserrat SemiBold"

	p := New(WithTheme(brand))
	p.AddSlide()
	files := generateFrom(t, p)
	theme := string(files["ppt/theme/theme1.xml"])

	for _, want := range []string{
		`name="Acme"`,
		`<a:srgbClr val="1F497D">`, // brand accent1
		`typeface="Montserrat"`,
		`typeface="Montserrat SemiBold"`,
	} {
		if !strings.Contains(theme, want) {
			t.Errorf("expected themed part to contain %q, got:\n%s", want, theme)
		}
	}
	// The default Office accent1 must be gone once overridden.
	if strings.Contains(theme, `<a:srgbClr val="4472C4">`) {
		t.Errorf("expected the default accent1 (4472C4) to be replaced by the brand color, got:\n%s", theme)
	}
}

func TestNew_WithoutThemeOptionKeepsDefaultOfficeTheme(t *testing.T) {
	p := New()
	p.AddSlide()
	files := generateFrom(t, p)
	theme := string(files["ppt/theme/theme1.xml"])

	if !strings.Contains(theme, `<a:srgbClr val="4472C4">`) {
		t.Errorf("expected the default Office accent1 (4472C4) with no WithTheme, got:\n%s", theme)
	}
}

func TestWithTheme_SchemeColorReferencesResolveToBrandPalette(t *testing.T) {
	// The whole point of theming: a shape references a color by slot
	// (accent1), and the brand color lives once in the theme part — so the
	// slide stays slot-referenced while the theme carries the brand hex.
	brand := DefaultTheme()
	brand.Colors.Accent1 = drawingml.Color{R: 0x1F, G: 0x49, B: 0x7D}

	p := New(WithTheme(brand))
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).FillScheme(SchemeAccent1)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])
	theme := string(files["ppt/theme/theme1.xml"])

	// The slide references the slot, not the hex — so the SAME shape recolors
	// when the theme changes.
	if !strings.Contains(slide, `<a:schemeClr val="accent1">`) {
		t.Errorf("expected the shape to reference accent1 by slot, got:\n%s", slide)
	}
	if strings.Contains(slide, "1F497D") {
		t.Errorf("expected the brand hex to live in the theme, not be inlined on the slide, got:\n%s", slide)
	}
	if !strings.Contains(theme, `<a:srgbClr val="1F497D">`) {
		t.Errorf("expected the brand accent1 hex in the theme part, got:\n%s", theme)
	}
}

func TestRenderThemeXML_EscapesThemeName(t *testing.T) {
	brand := DefaultTheme()
	brand.Name = `Ben & "Jerry" <Co>`

	xmlBytes := renderThemeXML(brand)

	var probe any
	if err := xml.Unmarshal(xmlBytes, &probe); err != nil {
		t.Fatalf("theme with reserved chars in Name is not well-formed XML: %v", err)
	}
	if !strings.Contains(string(xmlBytes), "&amp;") {
		t.Errorf("expected the ampersand in the theme name to be escaped, got:\n%s", string(xmlBytes))
	}
}

func TestWithTheme_PartialThemeFillsUnsetSlotsFromDefault(t *testing.T) {
	// Overriding only Accent1 must NOT render the other eleven slots as black
	// (the zero drawingml.Color) — they fall back to the default Office palette.
	brand := Theme{Colors: ThemeColors{Accent1: drawingml.Color{R: 0x1F, G: 0x49, B: 0x7D}}}

	p := New(WithTheme(brand))
	p.AddSlide()
	theme := string(generateFrom(t, p)["ppt/theme/theme1.xml"])

	if !strings.Contains(theme, `<a:srgbClr val="1F497D">`) {
		t.Errorf("expected the overridden accent1, got:\n%s", theme)
	}
	// Unset slots keep the Office defaults, not 000000.
	for _, want := range []string{
		`<a:srgbClr val="ED7D31">`,                 // accent2 default
		`<a:srgbClr val="70AD47">`,                 // accent6 default
		`<a:srgbClr val="E7E6E6">`,                 // lt2 default
		`<a:sysClr val="window" lastClr="FFFFFF">`, // lt1 stays system white, not black
		`typeface="Calibri"`,                       // fonts still default
	} {
		if !strings.Contains(theme, want) {
			t.Errorf("expected unset slot to keep its default (%q), got:\n%s", want, theme)
		}
	}
	// No color SLOT collapsed to black. (Check only the clrScheme fragment —
	// the fixed fmtScheme carries a legitimate 000000 shadow color.)
	start := strings.Index(theme, "<a:clrScheme")
	end := strings.Index(theme, "</a:clrScheme>")
	if start == -1 || end == -1 {
		t.Fatalf("could not locate the clrScheme fragment")
	}
	if strings.Contains(theme[start:end], `<a:srgbClr val="000000">`) {
		t.Errorf("expected no black color slots from a partial theme, got:\n%s", theme[start:end])
	}
}

func TestDefaultTheme_Dk1Lt1UseSystemColorsForAccessibility(t *testing.T) {
	// dk1/lt1 must be sysClr (windowText/window), like Office's own theme, so
	// text/background follow the viewer's OS (High Contrast, etc.) — not
	// hardcoded srgbClr black/white.
	s := string(renderThemeXML(DefaultTheme()))
	if !strings.Contains(s, `<a:sysClr val="windowText" lastClr="000000">`) {
		t.Errorf("expected dk1 as sysClr windowText, got:\n%s", s)
	}
	if !strings.Contains(s, `<a:sysClr val="window" lastClr="FFFFFF">`) {
		t.Errorf("expected lt1 as sysClr window, got:\n%s", s)
	}
	// A custom Dark1 becomes the lastClr fallback while staying system-adaptive.
	custom := DefaultTheme()
	custom.Colors.Dark1 = drawingml.Color{R: 0x11, G: 0x22, B: 0x33}
	cs := string(renderThemeXML(custom))
	if !strings.Contains(cs, `<a:sysClr val="windowText" lastClr="112233">`) {
		t.Errorf("expected a custom Dark1 as the sysClr lastClr, got:\n%s", cs)
	}
}

func TestDefaultTheme_EmptyFontsFallBackToOffice(t *testing.T) {
	// A caller building a Theme{} literal (rather than starting from
	// DefaultTheme) leaves fonts empty; render must still emit valid, non-empty
	// typefaces rather than typeface="" for the major/minor Latin fonts.
	bare := Theme{Colors: DefaultTheme().Colors} // no fonts set
	s := string(renderThemeXML(bare))

	if !strings.Contains(s, `typeface="Calibri Light"`) || !strings.Contains(s, `typeface="Calibri"`) {
		t.Errorf("expected empty fonts to fall back to Office defaults, got:\n%s", s)
	}
}
