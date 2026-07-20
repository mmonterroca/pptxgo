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

// Package themes provides ready-to-use brand palettes for pptxgo, as
// pptx.Theme values you pass to pptx.New via pptx.WithTheme:
//
//	p := pptx.New(pptx.WithTheme(themes.Corporate()))
//
// Each preset is returned by a function (rather than a shared variable) so a
// caller can tweak a copy without disturbing the preset:
//
//	t := themes.Corporate()
//	t.Colors.Accent1 = pptx.RGB(0x1F, 0x49, 0x7D) // company navy
//	p := pptx.New(pptx.WithTheme(t))
//
// To build a wholly custom brand theme from scratch, start from
// pptx.DefaultTheme() and set the slots you care about — the twelve color
// slots (Dark1/Light1/Dark2/Light2, Accent1-6, Hyperlink/FollowedHyperlink)
// and the two fonts (Major for headings, Minor for body).
package themes

import "github.com/mmonterroca/pptxgo/pptx"

// Office returns pptxgo's built-in default — the standard Microsoft Office
// palette and typography (Calibri Light / Calibri). Identical to what
// pptx.New uses with no WithTheme option; exposed here for symmetry with the
// other presets. See pptx.DefaultTheme.
func Office() pptx.Theme {
	return pptx.DefaultTheme()
}

// Corporate returns a professional business palette: navy blue primary with a
// red accent, on white — for reports, proposals, and corporate decks.
func Corporate() pptx.Theme {
	return pptx.Theme{
		Name: "Corporate",
		Colors: pptx.ThemeColors{
			Dark1:             pptx.RGB(0x00, 0x00, 0x00), // text
			Light1:            pptx.RGB(0xFF, 0xFF, 0xFF), // background
			Dark2:             pptx.RGB(0x2F, 0x54, 0x96), // navy
			Light2:            pptx.RGB(0xD9, 0xD9, 0xD9), // light gray
			Accent1:           pptx.RGB(0x2F, 0x54, 0x96), // navy (primary)
			Accent2:           pptx.RGB(0x4F, 0x81, 0xBD), // light blue (secondary)
			Accent3:           pptx.RGB(0xC0, 0x00, 0x00), // red (accent/emphasis)
			Accent4:           pptx.RGB(0xFF, 0xC0, 0x00), // amber (warning)
			Accent5:           pptx.RGB(0x00, 0xB0, 0x50), // green (success)
			Accent6:           pptx.RGB(0x59, 0x59, 0x59), // dark gray (muted)
			Hyperlink:         pptx.RGB(0x2F, 0x54, 0x96),
			FollowedHyperlink: pptx.RGB(0x95, 0x4F, 0x72),
		},
		Fonts: pptx.ThemeFonts{Major: "Calibri", Minor: "Calibri"},
	}
}

// Modern returns a clean, minimalist palette: slate and a bright blue accent,
// with generous neutrals — for technical decks and white-paper-style content.
func Modern() pptx.Theme {
	return pptx.Theme{
		Name: "Modern",
		Colors: pptx.ThemeColors{
			Dark1:             pptx.RGB(0x2C, 0x3E, 0x50), // midnight blue (text)
			Light1:            pptx.RGB(0xFF, 0xFF, 0xFF), // background
			Dark2:             pptx.RGB(0x34, 0x49, 0x5E), // wet asphalt
			Light2:            pptx.RGB(0xEC, 0xF0, 0xF1), // clouds
			Accent1:           pptx.RGB(0x29, 0x80, 0xB9), // peter river blue (primary)
			Accent2:           pptx.RGB(0x95, 0xA5, 0xA6), // concrete (secondary)
			Accent3:           pptx.RGB(0x34, 0x49, 0x5E), // wet asphalt
			Accent4:           pptx.RGB(0xF3, 0x9C, 0x12), // orange (warning)
			Accent5:           pptx.RGB(0x27, 0xAE, 0x60), // nephritis (success)
			Accent6:           pptx.RGB(0xC0, 0x39, 0x2B), // pomegranate (error)
			Hyperlink:         pptx.RGB(0x29, 0x80, 0xB9),
			FollowedHyperlink: pptx.RGB(0x8E, 0x44, 0xAD),
		},
		Fonts: pptx.ThemeFonts{Major: "Segoe UI Light", Minor: "Segoe UI"},
	}
}
