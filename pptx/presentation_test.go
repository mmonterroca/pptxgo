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
	"archive/zip"
	"bytes"
	"encoding/xml"
	"regexp"
	"strings"
	"testing"
)

// This is the Go-level regression suite. It cannot substitute for the
// OpenXML SDK validator or an actual OOXML consumer (see PptxValidator/ and
// the manual open-in-PowerPoint step) — a document can satisfy every check
// below and still be schema-invalid. What it catches cheaply and on every
// `go test`, without any external tool, is exactly the class of bug that
// caused docxgo's own round-trip incidents: a relationship pointing at a
// part that doesn't exist, or a part nobody points at.

func generate(t *testing.T) map[string][]byte {
	t.Helper()
	var buf bytes.Buffer
	if err := New().Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	files := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open %s: %v", f.Name, err)
		}
		var b bytes.Buffer
		if _, err := b.ReadFrom(rc); err != nil {
			t.Fatalf("read %s: %v", f.Name, err)
		}
		rc.Close()
		files[f.Name] = b.Bytes()
	}
	return files
}

func TestNew_ProducesEveryRequiredPart(t *testing.T) {
	files := generate(t)

	for _, want := range []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels",
		"ppt/theme/theme1.xml",
		"ppt/slideMasters/slideMaster1.xml",
		"ppt/slideMasters/_rels/slideMaster1.xml.rels",
		"ppt/slideLayouts/slideLayout1.xml",
		"ppt/slideLayouts/_rels/slideLayout1.xml.rels",
		"ppt/slides/slide1.xml",
		"ppt/slides/_rels/slide1.xml.rels",
		"docProps/core.xml",
		"docProps/app.xml",
	} {
		if _, ok := files[want]; !ok {
			t.Errorf("missing required part %s", want)
		}
	}
}

func TestNew_EveryPartIsWellFormedXML(t *testing.T) {
	for name, data := range generate(t) {
		if !strings.HasSuffix(name, ".xml") && !strings.HasSuffix(name, ".rels") {
			continue
		}
		var v any
		if err := xml.Unmarshal(data, &v); err != nil {
			t.Errorf("%s is not well-formed XML: %v", name, err)
		}
	}
}

var relTargetRe = regexp.MustCompile(`Target="([^"]+)"\s*(?:TargetMode="([^"]+)"\s*)?/?>`)

// resolveRelTarget resolves a relationship's Target attribute against the
// part that owns the .rels file it came from, following the OPC convention
// (relative to the owning part's directory, not the package root).
func resolveRelTarget(relsPath, target string) string {
	// relsPath is ".../_rels/<owner-basename>.rels" or "_rels/.rels".
	dir := strings.TrimSuffix(relsPath, "/_rels/"+lastSegment(relsPath))
	if dir == relsPath { // root: "_rels/.rels"
		dir = ""
	}
	segments := strings.Split(dir, "/")
	if dir == "" {
		segments = nil
	}
	for _, part := range strings.Split(target, "/") {
		switch part {
		case ".":
			// no-op
		case "..":
			if len(segments) > 0 {
				segments = segments[:len(segments)-1]
			}
		default:
			segments = append(segments, part)
		}
	}
	return strings.Join(segments, "/")
}

func lastSegment(p string) string {
	i := strings.LastIndex(p, "/")
	return p[i+1:]
}

func TestNew_EveryInternalRelationshipTargetExists(t *testing.T) {
	files := generate(t)

	for name, data := range files {
		if !strings.HasSuffix(name, ".rels") {
			continue
		}
		for _, m := range relTargetRe.FindAllStringSubmatch(string(data), -1) {
			target, mode := m[1], m[2]
			if mode == "External" {
				continue
			}
			resolved := resolveRelTarget(name, target)
			if _, ok := files[resolved]; !ok {
				t.Errorf("%s references target %q, resolved to %q, which does not exist as a part",
					name, target, resolved)
			}
		}
	}
}

func TestNew_ContentTypesCoversEveryNonMediaPart(t *testing.T) {
	files := generate(t)
	ct := string(files["[Content_Types].xml"])

	for name := range files {
		if strings.HasSuffix(name, ".rels") || name == "[Content_Types].xml" {
			continue
		}
		if !strings.Contains(ct, `PartName="/`+name+`"`) {
			t.Errorf("[Content_Types].xml has no Override for part %s", name)
		}
	}
}

func TestNew_RelationshipIDsAreUniquePerOwner(t *testing.T) {
	idRe := regexp.MustCompile(`Id="([^"]+)"`)
	for name, data := range generate(t) {
		if !strings.HasSuffix(name, ".rels") {
			continue
		}
		seen := make(map[string]bool)
		for _, m := range idRe.FindAllStringSubmatch(string(data), -1) {
			id := m[1]
			if seen[id] {
				t.Errorf("%s has a duplicate relationship ID %s", name, id)
			}
			seen[id] = true
		}
	}
}
