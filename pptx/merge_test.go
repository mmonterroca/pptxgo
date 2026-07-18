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
	"bytes"
	"strings"
	"testing"
)

func TestTemplate_PlaceholderNamesFindsAllDistinctKeys(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	names, err := tmpl.PlaceholderNames()
	if err != nil {
		t.Fatalf("PlaceholderNames: %v", err)
	}
	want := []string{"client_name", "contact_email"}
	if len(names) != len(want) {
		t.Fatalf("PlaceholderNames() = %v, want %v", names, want)
	}
	for i := range want {
		if names[i] != want[i] {
			t.Errorf("PlaceholderNames()[%d] = %q, want %q", i, names[i], want[i])
		}
	}
}

func TestTemplate_MergeSubstitutesAcrossAllSlidesAndReportsCount(t *testing.T) {
	tmpl := openFixture(t, testdataSample)

	n, err := tmpl.Merge(MergeData{
		"client_name":   "Acme Corp",
		"contact_email": "sales@acme.example",
	})
	if err != nil {
		t.Fatalf("Merge: %v", err)
	}
	if n != 3 {
		t.Errorf("Merge() = %d substitutions, want 3 (title + 2 body placeholders)", n)
	}

	s1, err := tmpl.Slide(1)
	if err != nil {
		t.Fatalf("Slide(1): %v", err)
	}
	text, err := s1.Text()
	if err != nil {
		t.Fatalf("Text: %v", err)
	}
	if strings.Contains(text, "{{") {
		t.Errorf("expected no unmerged placeholders left, got %q", text)
	}
	for _, want := range []string{"Acme Corp Quarterly Review", "Prepared for Acme Corp", "Contact: sales@acme.example"} {
		if !strings.Contains(text, want) {
			t.Errorf("expected merged text to contain %q, got %q", want, text)
		}
	}
}

func TestTemplate_MergeLeavesUnmatchedPlaceholderInPlaceWithoutStrictMode(t *testing.T) {
	tmpl := openFixture(t, testdataSample)

	n, err := tmpl.Merge(MergeData{"client_name": "Acme Corp"}) // contact_email intentionally omitted
	if err != nil {
		t.Fatalf("Merge without strict mode: %v", err)
	}
	if n != 2 {
		t.Errorf("Merge() = %d, want 2 (only client_name's 2 occurrences)", n)
	}

	s1, _ := tmpl.Slide(1)
	text, _ := s1.Text()
	if !strings.Contains(text, "{{contact_email}}") {
		t.Errorf("expected the unmatched placeholder to survive untouched, got %q", text)
	}
}

func TestTemplate_MergeStrictModeErrorsOnUnmatchedPlaceholder(t *testing.T) {
	tmpl := openFixture(t, testdataSample)

	_, err := tmpl.Merge(MergeData{"client_name": "Acme Corp"}, WithStrictMode())
	if err == nil {
		t.Fatal("expected an error in strict mode for the unmatched contact_email placeholder")
	}
	if !strings.Contains(err.Error(), "contact_email") {
		t.Errorf("expected the error to name the missing key, got %v", err)
	}
}

func TestTemplate_MergeWithCustomDelimiters(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	// The fixture uses "{{key}}"; with different delimiters configured,
	// Merge must find nothing to substitute (0, not an error).
	n, err := tmpl.Merge(MergeData{"client_name": "Acme Corp"}, WithDelimiters("[[", "]]"))
	if err != nil {
		t.Fatalf("Merge with custom delimiters: %v", err)
	}
	if n != 0 {
		t.Errorf("Merge() with non-matching delimiters = %d, want 0", n)
	}
}

func TestOpenSlide_MergeOnlyAffectsThatSlide(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	s1, err := tmpl.Slide(1)
	if err != nil {
		t.Fatalf("Slide(1): %v", err)
	}

	n, err := s1.Merge(MergeData{"client_name": "Acme Corp", "contact_email": "sales@acme.example"})
	if err != nil {
		t.Fatalf("Slide.Merge: %v", err)
	}
	if n != 3 {
		t.Errorf("Slide.Merge() = %d, want 3", n)
	}

	s2, err := tmpl.Slide(2)
	if err != nil {
		t.Fatalf("Slide(2): %v", err)
	}
	text2, err := s2.Text()
	if err != nil {
		t.Fatalf("Slide(2).Text(): %v", err)
	}
	if text2 != "Thank you" {
		t.Errorf("expected slide 2 untouched by slide 1's merge, got %q", text2)
	}
}

func TestTemplate_ReplaceLiteralSubstring(t *testing.T) {
	tmpl := openFixture(t, testdataSample)

	n, err := tmpl.Replace("Thank you", "Gracias")
	if err != nil {
		t.Fatalf("Replace: %v", err)
	}
	if n != 1 {
		t.Errorf("Replace() = %d, want 1", n)
	}

	s2, _ := tmpl.Slide(2)
	text, _ := s2.Text()
	if text != "Gracias" {
		t.Errorf("Slide(2).Text() = %q, want %q", text, "Gracias")
	}
}

func TestTemplate_ReplaceNoMatchReturnsZero(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	n, err := tmpl.Replace("nonexistent-string", "x")
	if err != nil {
		t.Fatalf("Replace: %v", err)
	}
	if n != 0 {
		t.Errorf("Replace() = %d, want 0", n)
	}
}

func TestTemplate_MergeThenSaveThenReopenPureGo(t *testing.T) {
	// Round-trip coverage without dotnet: open -> merge -> save -> reopen,
	// asserting no error, placeholders gone, expected text present. This
	// runs everywhere go test runs, unlike the OpenXML SDK validation
	// wired into make check (which needs the dotnet toolchain).
	tmpl := openFixture(t, testdataSample)

	if _, err := tmpl.Merge(MergeData{
		"client_name":   "Acme Corp",
		"contact_email": "sales@acme.example",
	}); err != nil {
		t.Fatalf("Merge: %v", err)
	}

	var buf bytes.Buffer
	if err := tmpl.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	reopened, err := OpenFromBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("reopen merged+saved bytes: %v", err)
	}
	if reopened.SlideCount() != 2 {
		t.Fatalf("reopened SlideCount() = %d, want 2", reopened.SlideCount())
	}

	s1, err := reopened.Slide(1)
	if err != nil {
		t.Fatalf("reopened Slide(1): %v", err)
	}
	text, err := s1.Text()
	if err != nil {
		t.Fatalf("reopened Slide(1).Text(): %v", err)
	}
	if strings.Contains(text, "{{") {
		t.Errorf("expected no leftover placeholders after a merge+save+reopen round-trip, got %q", text)
	}
	if !strings.Contains(text, "Acme Corp") || !strings.Contains(text, "sales@acme.example") {
		t.Errorf("expected merged values to survive a save+reopen round-trip, got %q", text)
	}
}
