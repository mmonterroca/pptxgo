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
	"encoding/xml"
	"io"

	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// Write serializes the package to w as a ZIP archive: [Content_Types].xml
// (always derived from the current part set, never hand-maintained),
// _rels/.rels, every part's own .rels if it has relationships, and every
// part itself — written verbatim if Raw is set, marshaled from Value
// otherwise. This is the only path through which a package is written,
// whether every part was generated from scratch or loaded from a template
// and partially replaced.
func (p *Package) Write(w io.Writer) error {
	zw := zip.NewWriter(w)

	if err := writeXMLPart(zw, PathContentTypes, p.contentTypesXML()); err != nil {
		return errors.Wrap(err, "Package.Write: content types")
	}

	if err := writeRelsIfPresent(zw, PathRootRels, p.Relationships("")); err != nil {
		return errors.Wrap(err, "Package.Write: root relationships")
	}

	for _, part := range p.Parts() {
		if part.Raw != nil {
			if err := writeRawPart(zw, part.Path, part.Raw); err != nil {
				return errors.Wrap(err, "Package.Write: part "+part.Path)
			}
		} else {
			if err := writeXMLPart(zw, part.Path, part.Value); err != nil {
				return errors.Wrap(err, "Package.Write: part "+part.Path)
			}
		}

		if p.hasRelationships(part.Path) {
			relsPath := relsPathForPart(part.Path)
			if err := writeRelsIfPresent(zw, relsPath, p.Relationships(part.Path)); err != nil {
				return errors.Wrap(err, "Package.Write: relationships for "+part.Path)
			}
		}
	}

	return zw.Close()
}

func writeRelsIfPresent(zw *zip.Writer, path string, rm *RelationshipManager) error {
	if rm.Count() == 0 && path != PathRootRels {
		return nil
	}
	return writeXMLPart(zw, path, rm.ToXML())
}

// writeXMLPart marshals v with an XML declaration and writes it to path.
func writeXMLPart(zw *zip.Writer, path string, v any) error {
	w, err := zw.Create(path)
	if err != nil {
		return err
	}
	if _, err := w.Write([]byte(xml.Header)); err != nil {
		return err
	}
	enc := xml.NewEncoder(w)
	enc.Indent("", "  ")
	return enc.Encode(v)
}

// writeRawPart writes data verbatim to path.
func writeRawPart(zw *zip.Writer, path string, data []byte) error {
	w, err := zw.Create(path)
	if err != nil {
		return err
	}
	_, err = w.Write(data)
	return err
}
