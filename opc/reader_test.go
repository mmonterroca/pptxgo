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
	"os"
	"strings"
	"testing"
)

// testdataFixture is a real python-pptx-generated .pptx (a foreign
// producer, never pptxgo's own output) — using pptxgo's own output here
// would mask exactly the round-trip bugs this file exists to catch: a
// generate-then-read cycle never exercises content this package didn't
// already know how to produce.
const testdataFixture = "testdata/sample.pptx"

func readFixture(t *testing.T) []byte {
	t.Helper()
	data, err := os.ReadFile(testdataFixture)
	if err != nil {
		t.Fatalf("read fixture: %v", err)
	}
	return data
}

func TestOpenBytes_NullRoundTripPreservesContentPartsByteIdentical(t *testing.T) {
	original := readFixture(t)

	pkg, err := OpenBytes(original)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	var buf bytes.Buffer
	if err := pkg.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}

	origFiles := zipFiles(t, original)
	rtFiles := zipFiles(t, buf.Bytes())

	if len(origFiles) != len(rtFiles) {
		t.Errorf("expected the same part count across a null round-trip, got %d vs %d", len(origFiles), len(rtFiles))
	}

	for name, data := range origFiles {
		// Content_Types/.rels are always REGENERATED from in-memory state
		// (see Package.Write's own doc comment) — semantically equivalent,
		// never byte-identical to the original. Every other part was
		// loaded via AddRawPart and must come back out exactly as loaded.
		if name == PathContentTypes || strings.HasSuffix(name, ".rels") {
			continue
		}
		rtData, ok := rtFiles[name]
		if !ok {
			t.Errorf("round-tripped package is missing part %s", name)
			continue
		}
		if data != rtData {
			t.Errorf("part %s changed across a null round-trip (no edits were made)", name)
		}
	}
}

func TestOpenBytes_NullRoundTripReopensCleanly(t *testing.T) {
	original := readFixture(t)

	pkg, err := OpenBytes(original)
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}
	var buf bytes.Buffer
	if err := pkg.Write(&buf); err != nil {
		t.Fatalf("Write: %v", err)
	}

	pkg2, err := OpenBytes(buf.Bytes())
	if err != nil {
		t.Fatalf("reopen round-tripped bytes: %v", err)
	}
	if len(pkg2.Parts()) != len(pkg.Parts()) {
		t.Errorf("expected the same part count after reopening the round-tripped bytes, got %d vs %d",
			len(pkg2.Parts()), len(pkg.Parts()))
	}
}

func TestOpenBytes_PreservesRelationships(t *testing.T) {
	pkg, err := OpenBytes(readFixture(t))
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	rootRels := pkg.Relationships("").All()
	if len(rootRels) == 0 {
		t.Fatal("expected the root package to have relationships (officeDocument, core properties, ...)")
	}

	var sawPresentationRel bool
	for _, r := range rootRels {
		if strings.Contains(r.Target, "presentation.xml") {
			sawPresentationRel = true
		}
	}
	if !sawPresentationRel {
		t.Errorf("expected a root relationship targeting presentation.xml, got %+v", rootRels)
	}

	presRels := pkg.Relationships("ppt/presentation.xml").All()
	var sawSlideRel bool
	for _, r := range presRels {
		if strings.Contains(r.Target, "slide") {
			sawSlideRel = true
		}
	}
	if !sawSlideRel {
		t.Errorf("expected presentation.xml to have at least one slide relationship, got %+v", presRels)
	}
}

func TestOpenBytes_FutureAddCallsDoNotCollideWithLoadedRelationshipIDs(t *testing.T) {
	pkg, err := OpenBytes(readFixture(t))
	if err != nil {
		t.Fatalf("OpenBytes: %v", err)
	}

	rm := pkg.Relationships("ppt/presentation.xml")
	existing := make(map[string]bool)
	for _, r := range rm.All() {
		existing[r.ID] = true
	}

	newID, err := rm.Add("http://example.com/rel/thing", "target.xml", "Internal")
	if err != nil {
		t.Fatalf("Add: %v", err)
	}
	if existing[newID] {
		t.Errorf("newly generated relationship ID %s collides with one loaded from the fixture", newID)
	}
}

func buildTestZip(t *testing.T, entries map[string][]byte) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for name, data := range entries {
		w, err := zw.Create(name)
		if err != nil {
			t.Fatalf("zip.Create(%s): %v", name, err)
		}
		if _, err := w.Write(data); err != nil {
			t.Fatalf("write %s: %v", name, err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatalf("zip.Close: %v", err)
	}
	return buf.Bytes()
}

func TestOpenBytes_MissingContentTypesIsAnError(t *testing.T) {
	// A package with no [Content_Types].xml at all is not a valid OPC
	// package -- OpenBytes must reject it, not silently produce an empty
	// Package.
	data := buildTestZip(t, map[string][]byte{
		"ppt/presentation.xml": []byte("<x/>"),
	})

	if _, err := OpenBytes(data); err == nil {
		t.Fatal("expected an error for a package missing [Content_Types].xml")
	}
}

func TestOpenBytes_UndeclaredContentTypeIsAnError(t *testing.T) {
	data := buildTestZip(t, map[string][]byte{
		PathContentTypes:       []byte(`<?xml version="1.0"?><Types xmlns="` + NamespaceContentTypes + `"></Types>`),
		"ppt/presentation.xml": []byte("<x/>"),
	})

	if _, err := OpenBytes(data); err == nil {
		t.Fatal("expected an error for a part with no Override or Default content type")
	}
}

func TestPreallocSize_ClampsUntrustedUncompressedSize(t *testing.T) {
	cases := []struct {
		name string
		in   uint64
		want int
	}{
		{"zero", 0, 0},
		{"small", 1024, 1024},
		{"exactly max", maxPreallocBytes, maxPreallocBytes},
		{"just over max (a lying/bomb header)", maxPreallocBytes + 1, 0},
		{"max uint64 (would overflow int on 32-bit)", 1<<64 - 1, 0},
	}
	for _, c := range cases {
		if got := preallocSize(c.in); got != c.want {
			t.Errorf("%s: preallocSize(%d) = %d, want %d", c.name, c.in, got, c.want)
		}
	}
}

func TestOwnerPathFromRelsPath(t *testing.T) {
	cases := []struct {
		relsPath  string
		wantOwner string
		wantOK    bool
	}{
		{"_rels/.rels", "", true},
		{"ppt/_rels/presentation.xml.rels", "ppt/presentation.xml", true},
		{"ppt/slides/_rels/slide1.xml.rels", "ppt/slides/slide1.xml", true},
		{"ppt/slides/slide1.xml", "", false},       // not a .rels part
		{"ppt/notrels/slide1.xml.rels", "", false}, // not inside a _rels dir
	}
	for _, c := range cases {
		owner, ok := ownerPathFromRelsPath(c.relsPath)
		if ok != c.wantOK || owner != c.wantOwner {
			t.Errorf("ownerPathFromRelsPath(%q) = (%q, %v), want (%q, %v)", c.relsPath, owner, ok, c.wantOwner, c.wantOK)
		}
	}
}

func TestContentTypeForPart_OverrideWinsOverDefault(t *testing.T) {
	overrides := map[string]string{"ppt/presentation.xml": "application/vnd.example.presentation+xml"}
	defaults := map[string]string{"xml": ContentTypeXML}

	ct, isOverride, ok := contentTypeForPart("ppt/presentation.xml", overrides, defaults)
	if !ok || !isOverride || ct != "application/vnd.example.presentation+xml" {
		t.Errorf("expected the Override to win, got (%q, %v, %v)", ct, isOverride, ok)
	}

	ct, isOverride, ok = contentTypeForPart("ppt/theme/theme1.xml", overrides, defaults)
	if !ok || isOverride || ct != ContentTypeXML {
		t.Errorf("expected the Default to apply, got (%q, %v, %v)", ct, isOverride, ok)
	}

	if _, _, ok := contentTypeForPart("ppt/media/image1.png", overrides, defaults); ok {
		t.Error("expected no content type resolved for an undeclared extension")
	}
}
