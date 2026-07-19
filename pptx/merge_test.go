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

func TestTemplate_MergeWithBracketDelimitersDoesNotSpanAcrossAnUnclosedOne(t *testing.T) {
	// Regression: a naive "[^{}]+?" excluded-char class doesn't protect
	// custom delimiters at all. With "[["/"]]", an unclosed "[[key1" ahead
	// of a real "[[key2]]" must not be captured as one bogus span
	// ("key1 ... [[key2"); the closing "]]" must belong only to key2.
	tmpl := openFixture(t, testdataSample)
	s, err := tmpl.Slide(1)
	if err != nil {
		t.Fatalf("Slide(1): %v", err)
	}
	if _, err := s.Replace("{{client_name}}", "[[key1 leftover [[key2]]"); err != nil {
		t.Fatalf("Replace (test setup): %v", err)
	}

	names, err := tmpl.PlaceholderNames(WithDelimiters("[[", "]]"))
	if err != nil {
		t.Fatalf("PlaceholderNames: %v", err)
	}
	if len(names) != 1 || names[0] != "key2" {
		t.Errorf("PlaceholderNames() = %v, want exactly [\"key2\"] (not a span starting at key1)", names)
	}
}

func TestTemplate_MergeStrictModeDeduplicatesRepeatedMissingKey(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	// contact_email appears once on slide 1; add a second occurrence on
	// slide 2 so the same missing key would be recorded twice without
	// deduplication.
	s2, err := tmpl.Slide(2)
	if err != nil {
		t.Fatalf("Slide(2): %v", err)
	}
	if _, err := s2.Replace("Thank you", "Thanks {{contact_email}}"); err != nil {
		t.Fatalf("Replace (test setup): %v", err)
	}

	_, err = tmpl.Merge(MergeData{"client_name": "Acme Corp"}, WithStrictMode())
	if err == nil {
		t.Fatal("expected an error for the unmatched contact_email placeholder")
	}
	// errors.InvalidArgument's own formatting embeds the value once via
	// "value=..." and once again in this package's own message text, so
	// the key name legitimately appears twice in the full error string
	// even when correctly deduplicated — what a duplication BUG would
	// additionally produce is the key repeated back-to-back within the
	// comma-joined list itself ("contact_email, contact_email").
	if msg := err.Error(); strings.Contains(msg, "contact_email, contact_email") {
		t.Errorf("expected contact_email to appear only once in the joined missing-key list, got %q", msg)
	}
}

func TestTemplate_PlaceholderNamesWithCustomDelimiters(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	s, err := tmpl.Slide(1)
	if err != nil {
		t.Fatalf("Slide(1): %v", err)
	}
	if _, err := s.Replace("{{client_name}}", "[[account]]"); err != nil {
		t.Fatalf("Replace (test setup): %v", err)
	}

	// Default delimiters must not find the bracket-style token.
	defaultNames, err := tmpl.PlaceholderNames()
	if err != nil {
		t.Fatalf("PlaceholderNames: %v", err)
	}
	for _, n := range defaultNames {
		if n == "account" {
			t.Errorf("expected default delimiters not to match a [[account]] token, got %v", defaultNames)
		}
	}

	bracketNames, err := tmpl.PlaceholderNames(WithDelimiters("[[", "]]"))
	if err != nil {
		t.Fatalf("PlaceholderNames with custom delimiters: %v", err)
	}
	found := false
	for _, n := range bracketNames {
		if n == "account" {
			found = true
		}
	}
	if !found {
		t.Errorf("expected PlaceholderNames(WithDelimiters(\"[[\", \"]]\")) to find \"account\", got %v", bracketNames)
	}
}

func TestTemplate_PlaceholderNamesAgreesWithMergeOnFormatSplitPlaceholders(t *testing.T) {
	// A placeholder whose key is separately formatted lands in 3 runs with
	// differing rPr, which groupRuns keeps separate, so Merge cannot
	// substitute it. PlaceholderNames must NOT report it either (it runs
	// the same run-grouping) — otherwise strict Merge would look like it
	// succeeded while the literal "{{client_name}}" survived in the deck.
	splitXML := `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p>` +
		`<a:r><a:rPr lang="en-US"/><a:t>{{</a:t></a:r>` +
		`<a:r><a:rPr lang="en-US" b="1"/><a:t>client_name</a:t></a:r>` +
		`<a:r><a:rPr lang="en-US"/><a:t>}}</a:t></a:r>` +
		`</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`

	tmpl := openFixture(t, testdataSample)
	s, err := tmpl.Slide(1)
	if err != nil {
		t.Fatalf("Slide(1): %v", err)
	}
	if err := tmpl.writeSlideBytes(s.path, []byte(splitXML)); err != nil {
		t.Fatalf("writeSlideBytes (test setup): %v", err)
	}

	names, err := tmpl.PlaceholderNames()
	if err != nil {
		t.Fatalf("PlaceholderNames: %v", err)
	}
	for _, n := range names {
		if n == "client_name" {
			t.Errorf("PlaceholderNames reported a format-split placeholder Merge cannot substitute: %v", names)
		}
	}

	// And crucially: strict Merge must NOT falsely report success on that
	// same unsubstitutable placeholder — it's simply not a placeholder the
	// engine recognizes, so a strict Merge with no data for it succeeds
	// (there is nothing it claims to match), and the literal text remains.
	n, err := tmpl.Merge(MergeData{}, WithStrictMode())
	if err != nil {
		t.Fatalf("Merge: %v", err)
	}
	if n != 0 {
		t.Errorf("expected 0 substitutions for a format-split placeholder, got %d", n)
	}
	text, err := s.Text()
	if err != nil {
		t.Fatalf("Text: %v", err)
	}
	if !strings.Contains(text, "client_name") {
		t.Errorf("expected the un-substitutable placeholder text to survive, got %q", text)
	}
}

func TestTemplate_ReplaceEmptyOldIsAnError(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	if _, err := tmpl.Replace("", "X"); err == nil {
		t.Fatal("expected Replace(\"\", ...) to error instead of corrupting every run")
	}
	s, err := tmpl.Slide(1)
	if err != nil {
		t.Fatalf("Slide(1): %v", err)
	}
	if _, err := s.Replace("", "X"); err == nil {
		t.Fatal("expected OpenSlide.Replace(\"\", ...) to error")
	}
	// The deck must be untouched — no corruption from the rejected call.
	text, err := s.Text()
	if err != nil {
		t.Fatalf("Text: %v", err)
	}
	if !strings.Contains(text, "{{client_name}} Quarterly Review") {
		t.Errorf("expected the slide text intact after a rejected empty-old Replace, got %q", text)
	}
}

func TestTemplate_MergeWithEmptyDelimitersIsAnErrorNotAPanic(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	if _, err := tmpl.Merge(MergeData{"client_name": "Acme"}, WithDelimiters("", "")); err == nil {
		t.Fatal("expected empty delimiters to error instead of panicking")
	}
	if _, err := tmpl.PlaceholderNames(WithDelimiters("{{", "")); err == nil {
		t.Fatal("expected a half-empty delimiter pair to error")
	}
	s, err := tmpl.Slide(1)
	if err != nil {
		t.Fatalf("Slide(1): %v", err)
	}
	if _, err := s.Merge(MergeData{}, WithDelimiters("", "]]")); err == nil {
		t.Fatal("expected OpenSlide.Merge with an empty open delimiter to error")
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
