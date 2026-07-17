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

package opc

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"strings"
	"testing"
)

type stubValue struct {
	XMLName xml.Name `xml:"stub"`
	Text    string   `xml:"text,attr"`
}

func zipFiles(t *testing.T, data []byte) map[string]string {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}
	out := make(map[string]string, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open %s: %v", f.Name, err)
		}
		var buf bytes.Buffer
		if _, err := buf.ReadFrom(rc); err != nil {
			t.Fatalf("read %s: %v", f.Name, err)
		}
		rc.Close()
		out[f.Name] = buf.String()
	}
	return out
}

func TestPackage_GeneratedAndRawPartsShareOneWritePath(t *testing.T) {
	// The whole point of a part-centric package: a part built from scratch
	// (Value) and a part loaded verbatim from a template (Raw) are written
	// through the exact same code, with no "round-trip mode" branch.
	pkg := NewPackage()
	pkg.AddPart("ppt/generated.xml", "application/vnd.example.generated+xml", &stubValue{Text: "hello"})
	pkg.AddRawPart("ppt/preserved.xml", "application/vnd.example.preserved+xml", []byte(`<preserved/>`))

	var buf bytes.Buffer
	if err := pkg.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}

	files := zipFiles(t, buf.Bytes())

	generated, ok := files["ppt/generated.xml"]
	if !ok {
		t.Fatal("missing ppt/generated.xml")
	}
	if !strings.Contains(generated, `<stub text="hello">`) {
		t.Errorf("generated part not marshaled from Value: %s", generated)
	}

	preserved, ok := files["ppt/preserved.xml"]
	if !ok {
		t.Fatal("missing ppt/preserved.xml")
	}
	if preserved != `<preserved/>` {
		t.Errorf("raw part not written verbatim: %q", preserved)
	}
}

func TestPackage_ContentTypesDerivedFromParts(t *testing.T) {
	pkg := NewPackage()
	pkg.AddPart("ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml", &stubValue{})
	pkg.AddMediaPart("ppt/media/image1.png", ContentTypePNG, []byte{0x89, 'P', 'N', 'G'})
	pkg.AddMediaPart("ppt/media/image2.png", ContentTypePNG, []byte{0x89, 'P', 'N', 'G'})

	var buf bytes.Buffer
	if err := pkg.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}
	files := zipFiles(t, buf.Bytes())

	ct, ok := files[PathContentTypes]
	if !ok {
		t.Fatal("missing [Content_Types].xml")
	}

	// The presentation part is specific enough to need its own Override.
	if !strings.Contains(ct, `PartName="/ppt/presentation.xml"`) {
		t.Errorf("expected an Override for the presentation part, got: %s", ct)
	}
	// Media parts share a single Default by extension, not one Override each.
	if strings.Count(ct, `Extension="png"`) != 1 {
		t.Errorf("expected exactly one png Default, got: %s", ct)
	}
	if strings.Contains(ct, `PartName="/ppt/media/image1.png"`) {
		t.Errorf("media part should not receive its own Override: %s", ct)
	}

	if _, ok := files["ppt/media/image1.png"]; !ok {
		t.Error("missing ppt/media/image1.png")
	}
}

func TestPackage_PerPartRelationshipsAreScopedIndependently(t *testing.T) {
	// The bug class this guards against is real: docxgo shipped with rIds
	// resolved against the wrong scope until a dedicated fix (per-part
	// .rels parsing) landed nine months after the writer was first built.
	pkg := NewPackage()
	pkg.AddPart("ppt/slides/slide1.xml", "application/vnd.example.slide+xml", &stubValue{})
	pkg.AddPart("ppt/slides/slide2.xml", "application/vnd.example.slide+xml", &stubValue{})

	id1, err := pkg.Relationships("ppt/slides/slide1.xml").AddImage("../media/image1.png")
	if err != nil {
		t.Fatalf("AddImage slide1: %v", err)
	}
	id2, err := pkg.Relationships("ppt/slides/slide2.xml").AddImage("../media/image2.png")
	if err != nil {
		t.Fatalf("AddImage slide2: %v", err)
	}
	if id1 != id2 {
		t.Fatalf("expected both slides to independently start at the same first rId, got %s and %s", id1, id2)
	}

	var buf bytes.Buffer
	if err := pkg.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}
	files := zipFiles(t, buf.Bytes())

	rels1, ok := files["ppt/slides/_rels/slide1.xml.rels"]
	if !ok {
		t.Fatal("missing ppt/slides/_rels/slide1.xml.rels")
	}
	if !strings.Contains(rels1, "image1.png") {
		t.Errorf("slide1 rels missing its own target: %s", rels1)
	}
	if strings.Contains(rels1, "image2.png") {
		t.Errorf("slide1 rels leaked slide2's relationship: %s", rels1)
	}

	rels2, ok := files["ppt/slides/_rels/slide2.xml.rels"]
	if !ok {
		t.Fatal("missing ppt/slides/_rels/slide2.xml.rels")
	}
	if !strings.Contains(rels2, "image2.png") {
		t.Errorf("slide2 rels missing its own target: %s", rels2)
	}
}

func TestPackage_RootRelsAlwaysWritten(t *testing.T) {
	pkg := NewPackage()
	pkg.AddPart("ppt/presentation.xml", "application/vnd.example.presentation+xml", &stubValue{})
	if _, err := pkg.Relationships("").Add("http://example.com/rel/officeDocument", "ppt/presentation.xml", "Internal"); err != nil {
		t.Fatalf("Add root rel: %v", err)
	}

	var buf bytes.Buffer
	if err := pkg.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}
	files := zipFiles(t, buf.Bytes())

	if _, ok := files[PathRootRels]; !ok {
		t.Fatal("missing _rels/.rels")
	}
}

func TestPackage_WriteIsByteDeterministic(t *testing.T) {
	// Regression for the map-iteration-order bug: a package with several
	// relationships per owner must serialize to identical bytes across runs,
	// or golden-file tests and reproducible builds break. Relationships are
	// stored in a map, so without an explicit ordering this fails
	// intermittently (Go randomizes map iteration).
	build := func() []byte {
		pkg := NewPackage()
		pkg.AddPart("ppt/presentation.xml", "application/vnd.example.presentation+xml", &stubValue{})
		rm := pkg.Relationships("ppt/presentation.xml")
		for i := 0; i < 12; i++ {
			if _, err := rm.Add("http://example.com/rel/thing", "target"+string(rune('a'+i))+".xml", "Internal"); err != nil {
				t.Fatalf("Add: %v", err)
			}
		}
		var buf bytes.Buffer
		if err := pkg.Write(&buf); err != nil {
			t.Fatalf("Write: %v", err)
		}
		return buf.Bytes()
	}

	first := build()
	for i := 0; i < 5; i++ {
		if got := build(); !bytes.Equal(first, got) {
			t.Fatalf("Write produced different bytes across runs (run %d): relationship order is non-deterministic", i+1)
		}
	}
}

func TestPackage_ExtensionlessMediaStillGetsContentType(t *testing.T) {
	// An extensionless media path has no extension to hang a Default on; it
	// must still be declared via an Override, or Office rejects the package.
	pkg := NewPackage()
	pkg.AddMediaPart("ppt/media/image1", ContentTypePNG, []byte{0x89, 'P', 'N', 'G'})

	var buf bytes.Buffer
	if err := pkg.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}
	ct := zipFiles(t, buf.Bytes())[PathContentTypes]

	if !strings.Contains(ct, `PartName="/ppt/media/image1"`) {
		t.Errorf("extensionless media part has no content-type declaration: %s", ct)
	}
}

func TestPackage_AddMediaPartConflictingContentTypeGetsOverride(t *testing.T) {
	// Two media parts sharing an extension but declaring different content
	// types must not both hide behind the first one's Default — the second
	// needs its own Override, or Office resolves it to the wrong type.
	pkg := NewPackage()
	pkg.AddMediaPart("ppt/media/image1.png", ContentTypePNG, []byte{1})
	pkg.AddMediaPart("ppt/media/image2.png", "image/custom", []byte{2})

	var buf bytes.Buffer
	if err := pkg.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}
	ct := string(zipFiles(t, buf.Bytes())[PathContentTypes])

	if !strings.Contains(ct, `Extension="png" ContentType="image/png"`) {
		t.Errorf("expected the first-registered png Default to stay image/png: %s", ct)
	}
	if !strings.Contains(ct, `PartName="/ppt/media/image2.png" ContentType="image/custom"`) {
		t.Errorf("expected image2.png to get its own Override for the conflicting type: %s", ct)
	}
}

func TestNormalizePartPath_ConvertsBackslashes(t *testing.T) {
	if got, want := normalizePartPath(`ppt\media\image1.png`), "ppt/media/image1.png"; got != want {
		t.Errorf("normalizePartPath(backslash path) = %q, want %q", got, want)
	}
}

func TestPart_NotFoundVsFound(t *testing.T) {
	pkg := NewPackage()
	pkg.AddPart("ppt/presentation.xml", "application/vnd.example.presentation+xml", &stubValue{})

	if _, ok := pkg.Part("missing.xml"); ok {
		t.Error("expected missing.xml to be absent")
	}
	if _, ok := pkg.Part("ppt/presentation.xml"); !ok {
		t.Error("expected ppt/presentation.xml to be present")
	}
	// Leading-slash and bare paths must refer to the same part.
	if _, ok := pkg.Part("/ppt/presentation.xml"); !ok {
		t.Error("expected leading-slash lookup to normalize to the same part")
	}
}
