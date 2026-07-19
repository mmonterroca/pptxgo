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
	"io"
	"os"
	"path"
	"sort"
	"strings"

	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// Open reads an existing OPC package from disk into an in-memory Package,
// ready for Write to re-save (with or without further edits) — see
// OpenBytes for how every part is loaded.
func Open(pth string) (*Package, error) {
	data, err := os.ReadFile(pth)
	if err != nil {
		return nil, errors.Wrap(err, "opc.Open")
	}
	return OpenBytes(data)
}

// OpenReader reads an existing OPC package from r. The whole stream is
// buffered first — archive/zip requires an io.ReaderAt plus a known size,
// which an arbitrary io.Reader cannot provide directly.
func OpenReader(r io.Reader) (*Package, error) {
	data, err := io.ReadAll(r)
	if err != nil {
		return nil, errors.Wrap(err, "opc.OpenReader")
	}
	return OpenBytes(data)
}

// OpenBytes reads an existing OPC package already held in memory. Every
// part is hydrated through the same mutation surface a caller could use
// directly (AddRawPart, AddMediaPart, RegisterExisting, EnsureAtLeast) —
// this function adds no new way to mutate a Package, only a way to seed one
// from a real file. Content parts stay as Raw bytes; nothing is unmarshaled
// into a content-schema struct (opc knows nothing about PresentationML/
// WordprocessingML content and never will — see the package doc comment).
//
// [Content_Types].xml and every .rels part ARE parsed (their own structs,
// XMLContentTypes/XMLRelationships, round-trip cleanly — see their doc
// comments), since that's the metadata OpenBytes itself needs to reconstruct
// the Package's part/relationship state; everything else is preserved
// verbatim. Write later regenerates equivalent (not byte-identical)
// Content_Types/.rels output from that reconstructed state, the same as it
// always has for a package assembled from scratch.
func OpenBytes(data []byte) (*Package, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, errors.Wrap(err, "opc.OpenBytes: not a valid ZIP/OPC package")
	}
	return loadFromZip(zr)
}

// loadFromZip does the actual reconstruction: split the archive into
// [Content_Types].xml, .rels parts, and everything else; resolve each
// content part's declared content type; then seed a fresh Package.
func loadFromZip(zr *zip.Reader) (*Package, error) {
	var contentTypesRaw []byte
	relsData := make(map[string][]byte)
	contentData := make(map[string][]byte)
	var contentOrder []string

	for _, f := range zr.File {
		if f.FileInfo().IsDir() {
			continue
		}
		data, err := readZipFile(f)
		if err != nil {
			return nil, errors.Wrap(err, "opc.OpenBytes: read "+f.Name)
		}

		switch {
		case f.Name == PathContentTypes:
			contentTypesRaw = data
		case strings.HasSuffix(f.Name, ".rels"):
			relsData[f.Name] = data
		default:
			contentData[f.Name] = data
			contentOrder = append(contentOrder, f.Name)
		}
	}

	if contentTypesRaw == nil {
		return nil, errors.InvalidArgument("opc.OpenBytes", "path", PathContentTypes, "package is missing "+PathContentTypes)
	}
	var ct XMLContentTypes
	if err := xml.Unmarshal(contentTypesRaw, &ct); err != nil {
		return nil, errors.Wrap(err, "opc.OpenBytes: parse "+PathContentTypes)
	}

	overrides := make(map[string]string, len(ct.Overrides))
	for _, o := range ct.Overrides {
		overrides[normalizePartPath(o.PartName)] = o.ContentType
	}
	defaults := make(map[string]string, len(ct.Defaults))
	for _, d := range ct.Defaults {
		defaults[strings.ToLower(d.Extension)] = d.ContentType
	}

	pkg := NewPackage()

	for _, pth := range contentOrder {
		normPth := normalizePartPath(pth)
		contentType, isOverride, ok := contentTypeForPart(normPth, overrides, defaults)
		if !ok {
			return nil, errors.InvalidArgument("opc.OpenBytes", "path", pth,
				"no Override or Default content type declared for this part in "+PathContentTypes)
		}
		if isOverride {
			pkg.AddRawPart(pth, contentType, contentData[pth])
		} else {
			pkg.AddMediaPart(pth, contentType, contentData[pth])
		}
	}

	relsPaths := make([]string, 0, len(relsData))
	for pth := range relsData {
		relsPaths = append(relsPaths, pth)
	}
	sort.Strings(relsPaths)

	for _, relsPath := range relsPaths {
		owner, ok := ownerPathFromRelsPath(relsPath)
		if !ok {
			return nil, errors.InvalidArgument("opc.OpenBytes", "path", relsPath, "malformed .rels part path")
		}
		var xr XMLRelationships
		if err := xml.Unmarshal(relsData[relsPath], &xr); err != nil {
			return nil, errors.Wrap(err, "opc.OpenBytes: parse "+relsPath)
		}
		rm := pkg.Relationships(owner)
		for _, r := range xr.Relationships {
			if err := rm.RegisterExisting(r.ID, r.Type, r.Target, r.TargetMode); err != nil {
				return nil, errors.Wrap(err, "opc.OpenBytes: register relationship in "+relsPath)
			}
		}
	}

	return pkg, nil
}

// contentTypeForPart resolves pth's declared content type: an Override
// (keyed by exact part path) wins over a Default (keyed by lowercased
// extension), mirroring how [Content_Types].xml itself layers the two —
// see AddPart vs. AddMediaPart. The bool result reports whether pth had its
// own Override (true) or fell through to an extension Default (false), so
// the caller can replay the same choice AddRawPart/AddMediaPart would have
// made when this part was first added.
func contentTypeForPart(pth string, overrides, defaults map[string]string) (contentType string, isOverride, ok bool) {
	if ct, exists := overrides[pth]; exists {
		return ct, true, true
	}
	ext := strings.TrimPrefix(strings.ToLower(path.Ext(pth)), ".")
	if ct, exists := defaults[ext]; exists {
		return ct, false, true
	}
	return "", false, false
}

// ownerPathFromRelsPath inverts relsPathForPart: given a .rels part's own
// path (e.g. "ppt/slides/_rels/slide1.xml.rels" or the root "_rels/.rels"),
// returns the part path it describes relationships FOR (e.g.
// "ppt/slides/slide1.xml", or "" for the package root).
func ownerPathFromRelsPath(relsPath string) (owner string, ok bool) {
	const relsExt = ".rels"
	if !strings.HasSuffix(relsPath, relsExt) {
		return "", false
	}
	dir := path.Dir(relsPath)
	if path.Base(dir) != "_rels" {
		return "", false
	}
	parentDir := path.Dir(dir)
	base := strings.TrimSuffix(path.Base(relsPath), relsExt)
	if parentDir == "." {
		return base, true // base == "" for the root "_rels/.rels" case
	}
	return parentDir + "/" + base, true
}

func readZipFile(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	// The zip entry's own header already declares its uncompressed size,
	// so the destination buffer can be sized once up front instead of
	// growing (and re-copying) repeatedly as io.Copy fills it.
	var buf bytes.Buffer
	buf.Grow(int(f.UncompressedSize64))
	if _, err := io.Copy(&buf, rc); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
