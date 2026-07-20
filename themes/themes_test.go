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

package themes_test

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"testing"

	"github.com/mmonterroca/pptxgo/pptx"
	"github.com/mmonterroca/pptxgo/themes"
)

// themePart saves a presentation built with the given theme and returns its
// ppt/theme/theme1.xml bytes.
func themePart(t *testing.T, theme pptx.Theme) string {
	t.Helper()
	p := pptx.New(pptx.WithTheme(theme))
	p.AddSlide()

	var buf bytes.Buffer
	if err := p.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}
	for _, f := range zr.File {
		if f.Name != "ppt/theme/theme1.xml" {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open theme: %v", err)
		}
		defer rc.Close()
		var b bytes.Buffer
		if _, err := b.ReadFrom(rc); err != nil {
			t.Fatalf("read theme: %v", err)
		}
		return b.String()
	}
	t.Fatal("ppt/theme/theme1.xml not found in output")
	return ""
}

func TestPresets_ProduceWellFormedThemedPart(t *testing.T) {
	cases := []struct {
		name  string
		theme pptx.Theme
		want  string // an accent1 hex that must appear
	}{
		{"Office", themes.Office(), `<a:srgbClr val="4472C4">`},
		{"Corporate", themes.Corporate(), `<a:srgbClr val="2F5496">`},
		{"Modern", themes.Modern(), `<a:srgbClr val="2980B9">`},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			part := themePart(t, tc.theme)
			var probe any
			if err := xml.Unmarshal([]byte(part), &probe); err != nil {
				t.Fatalf("%s theme is not well-formed XML: %v", tc.name, err)
			}
			if !contains(part, tc.want) {
				t.Errorf("%s theme: expected accent1 %q, got:\n%s", tc.name, tc.want, part)
			}
		})
	}
}

func TestPreset_IsCopyNotSharedState(t *testing.T) {
	// Mutating a returned preset must not affect a subsequent call.
	a := themes.Corporate()
	a.Colors.Accent1 = pptx.RGB(0x00, 0x00, 0x00)
	b := themes.Corporate()
	if b.Colors.Accent1 == a.Colors.Accent1 {
		t.Errorf("expected each Corporate() call to return an independent value; mutation leaked")
	}
}

func contains(haystack, needle string) bool {
	return bytes.Contains([]byte(haystack), []byte(needle))
}
