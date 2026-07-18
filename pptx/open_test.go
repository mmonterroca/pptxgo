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
	"os"
	"strings"
	"testing"
)

// testdataSample and testdataReordered are real python-pptx-generated
// fixtures (foreign producer, never pptxgo's own output — see
// opc/reader_test.go's own fixture doc comment for why that matters).
// testdataReordered is the same 2-slide deck with its <p:sldId> entries
// physically reversed in presentation.xml, so tests against it can prove
// navigation follows sldIdLst order rather than slideN.xml filename order.
const (
	testdataSample    = "testdata/sample.pptx"
	testdataReordered = "testdata/reordered.pptx"
)

func openFixture(t *testing.T, pth string) *Template {
	t.Helper()
	tmpl, err := Open(pth)
	if err != nil {
		t.Fatalf("Open(%s): %v", pth, err)
	}
	return tmpl
}

func TestOpen_SlideCountMatchesFixture(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	if got := tmpl.SlideCount(); got != 2 {
		t.Errorf("SlideCount() = %d, want 2", got)
	}
}

func TestOpen_SlidesInPresentationOrderNotFilenameOrder(t *testing.T) {
	// The reordered fixture's sldIdLst lists slide2.xml ("Thank you")
	// before slide1.xml ("{{client_name}}..."), even though slide1.xml
	// sorts first by filename -- Slide(1) must return slide2's content.
	tmpl := openFixture(t, testdataReordered)

	first, err := tmpl.Slide(1)
	if err != nil {
		t.Fatalf("Slide(1): %v", err)
	}
	firstText, err := first.Text()
	if err != nil {
		t.Fatalf("Slide(1).Text(): %v", err)
	}
	if !strings.Contains(firstText, "Thank you") {
		t.Errorf("expected Slide(1) to be the presentation-order-first slide (\"Thank you\"), got %q", firstText)
	}

	second, err := tmpl.Slide(2)
	if err != nil {
		t.Fatalf("Slide(2): %v", err)
	}
	secondText, err := second.Text()
	if err != nil {
		t.Fatalf("Slide(2).Text(): %v", err)
	}
	if !strings.Contains(secondText, "client_name") {
		t.Errorf("expected Slide(2) to be the presentation-order-second slide, got %q", secondText)
	}
}

func TestOpen_SlideOutOfRangeIsAnError(t *testing.T) {
	tmpl := openFixture(t, testdataSample)

	if _, err := tmpl.Slide(0); err == nil {
		t.Error("expected an error for Slide(0)")
	}
	if _, err := tmpl.Slide(3); err == nil {
		t.Error("expected an error for Slide(3) on a 2-slide deck")
	}
}

func TestOpen_SlidesReturnsAllInOrderWithCorrectIndex(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	slides := tmpl.Slides()

	if len(slides) != 2 {
		t.Fatalf("Slides() returned %d slides, want 2", len(slides))
	}
	for i, s := range slides {
		if s.Index() != i+1 {
			t.Errorf("slides[%d].Index() = %d, want %d", i, s.Index(), i+1)
		}
	}
}

func TestOpenSlide_TextConcatenatesParagraphsAndFindsPlaceholderTokens(t *testing.T) {
	tmpl := openFixture(t, testdataSample)
	s, err := tmpl.Slide(1)
	if err != nil {
		t.Fatalf("Slide(1): %v", err)
	}
	text, err := s.Text()
	if err != nil {
		t.Fatalf("Text: %v", err)
	}

	for _, want := range []string{"{{client_name}} Quarterly Review", "Prepared for {{client_name}}", "Contact: {{contact_email}}"} {
		if !strings.Contains(text, want) {
			t.Errorf("expected Text() to contain %q, got %q", want, text)
		}
	}
}

func TestOpen_FromBytesAndFromReaderMatchOpen(t *testing.T) {
	data, err := os.ReadFile(testdataSample)
	if err != nil {
		t.Fatalf("read fixture: %v", err)
	}

	byPath := openFixture(t, testdataSample)

	byBytes, err := OpenFromBytes(data)
	if err != nil {
		t.Fatalf("OpenFromBytes: %v", err)
	}
	byReader, err := OpenFromReader(bytes.NewReader(data))
	if err != nil {
		t.Fatalf("OpenFromReader: %v", err)
	}

	for name, tmpl := range map[string]*Template{"OpenFromBytes": byBytes, "OpenFromReader": byReader} {
		if tmpl.SlideCount() != byPath.SlideCount() {
			t.Errorf("%s: SlideCount() = %d, want %d (matching Open)", name, tmpl.SlideCount(), byPath.SlideCount())
		}
	}
}

func TestTemplate_SaveWithNoEditsReopensWithSameSlideCount(t *testing.T) {
	tmpl := openFixture(t, testdataSample)

	var buf bytes.Buffer
	if err := tmpl.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	reopened, err := OpenFromBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("reopen saved bytes: %v", err)
	}
	if reopened.SlideCount() != tmpl.SlideCount() {
		t.Errorf("SlideCount() after a no-edit save+reopen = %d, want %d", reopened.SlideCount(), tmpl.SlideCount())
	}
}
