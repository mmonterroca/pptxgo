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
	"encoding/xml"
	"io"
	"path"
	"strings"

	"github.com/mmonterroca/pptxgo/opc"
	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// Template is a handle onto an existing .pptx opened via Open/
// OpenFromBytes/OpenFromReader, for template-style editing: enumerating
// slides, inspecting or replacing their text, and saving the result back
// out.
//
// Deliberately NOT a *Presentation: XMLPresentation and the other content-
// model structs (xml.go) cannot be unmarshaled from a foreign
// presentation.xml (see nav.go's own doc comment for why), so a rehydrated
// *Presentation would carry an empty/wrong pres field, and AddSlide (which
// mutates pres.SldIdLst directly) would then corrupt the file. Every part
// loaded through Open* instead stays as opaque Raw bytes inside the
// underlying opc.Package — Template only ever reads or byte-splices slide
// content (see substitute.go), never structurally edits the package.
type Template struct {
	pkg        *opc.Package
	slidePaths []string // slidePaths[i] is slide i+1's part path (1-based Slide/Slides)
}

// Open reads an existing .pptx from disk.
func Open(pth string) (*Template, error) {
	pkg, err := opc.Open(pth)
	if err != nil {
		return nil, errors.Wrap(err, "pptx.Open")
	}
	return newTemplate(pkg)
}

// OpenFromBytes reads an existing .pptx already held in memory.
func OpenFromBytes(data []byte) (*Template, error) {
	pkg, err := opc.OpenBytes(data)
	if err != nil {
		return nil, errors.Wrap(err, "pptx.OpenFromBytes")
	}
	return newTemplate(pkg)
}

// OpenFromReader reads an existing .pptx from r.
func OpenFromReader(r io.Reader) (*Template, error) {
	pkg, err := opc.OpenReader(r)
	if err != nil {
		return nil, errors.Wrap(err, "pptx.OpenFromReader")
	}
	return newTemplate(pkg)
}

// newTemplate resolves slide order from the opened package's own
// presentation.xml plus its relationships: rId -> target comes from
// presentation.xml's own relationships (filtered to RelTypeSlide, using
// the clean, already-parseable opc relationship structs); slide ORDER
// comes from p:sldIdLst inside presentation.xml itself, read through the
// dedicated read-safe struct in nav.go rather than the write-only
// XMLPresentation.
func newTemplate(pkg *opc.Package) (*Template, error) {
	presPart, ok := pkg.Part(PathPresentation)
	if !ok {
		return nil, errors.NotFound("pptx.Open", PathPresentation)
	}
	if presPart.Raw == nil {
		return nil, errors.InvalidArgument("pptx.Open", "presentation.xml", "", "expected raw bytes for an opened package's presentation part")
	}

	var nav xmlPresentationNav
	if err := xml.Unmarshal(presPart.Raw, &nav); err != nil {
		return nil, errors.Wrap(err, "pptx.Open: parse "+PathPresentation)
	}

	ridToPath := make(map[string]string)
	for _, r := range pkg.Relationships(PathPresentation).All() {
		if r.Type != RelTypeSlide {
			continue
		}
		// Relationship targets on presentation.xml are relative to its own
		// directory ("ppt/"), the same convention AddSlide's own
		// p.presRels.Add(RelTypeSlide, "slides/slideN.xml", ...) writes —
		// see presentation.go.
		ridToPath[r.ID] = path.Clean(path.Join(path.Dir(PathPresentation), r.Target))
	}

	var slidePaths []string
	if nav.SldIdLst != nil {
		slidePaths = make([]string, 0, len(nav.SldIdLst.Entries))
		for _, entry := range nav.SldIdLst.Entries {
			pth, ok := ridToPath[entry.RID]
			if !ok {
				return nil, errors.InvalidArgument("pptx.Open", "sldId", entry.RID, "no matching slide relationship in presentation.xml's own relationships")
			}
			slidePaths = append(slidePaths, pth)
		}
	}

	return &Template{pkg: pkg, slidePaths: slidePaths}, nil
}

// SlideCount returns the number of slides, in presentation order.
func (t *Template) SlideCount() int {
	return len(t.slidePaths)
}

// Slide returns a handle onto the nth slide, 1-indexed — matching how a
// presentation's slides are numbered everywhere else (e.g. "slide 2").
func (t *Template) Slide(n int) (*OpenSlide, error) {
	if n < 1 || n > len(t.slidePaths) {
		return nil, errors.InvalidArgument("Template.Slide", "n", n, "out of range for this presentation's slide count")
	}
	return &OpenSlide{tmpl: t, index: n, path: t.slidePaths[n-1]}, nil
}

// Slides returns every slide, in presentation order.
func (t *Template) Slides() []*OpenSlide {
	out := make([]*OpenSlide, len(t.slidePaths))
	for i := range t.slidePaths {
		out[i] = &OpenSlide{tmpl: t, index: i + 1, path: t.slidePaths[i]}
	}
	return out
}

// Save writes the presentation to w. Every part this Template never
// touched passes through verbatim (see opc.OpenBytes), so opening a
// template and saving it back out with no edits at all reproduces the
// original content, modulo Content_Types/.rels always being regenerated
// from in-memory state — semantically equivalent, not byte-identical, the
// same as any package this library assembles.
func (t *Template) Save(w io.Writer) error {
	return t.pkg.Write(w)
}

// rawSlideBytes returns the raw XML bytes of the slide at path pth, for
// OpenSlide's read/substitution methods.
func (t *Template) rawSlideBytes(pth string) ([]byte, error) {
	part, ok := t.pkg.Part(pth)
	if !ok || part.Raw == nil {
		return nil, errors.NotFound("Template", pth)
	}
	return part.Raw, nil
}

// OpenSlide is a handle onto one slide within an opened Template, returned
// by Template.Slide/Slides.
type OpenSlide struct {
	tmpl  *Template
	index int // 1-based
	path  string
}

// Index returns this slide's 1-based position in the presentation.
func (s *OpenSlide) Index() int {
	return s.index
}

// Text returns every run of text on the slide, in document order, one
// paragraph per line. Unlike Replace/Merge (substitute.go), this needs no
// run-consolidation: concatenating a:t content in document order
// reconstructs the right text even when PowerPoint has split it across
// several <a:r> runs (autocorrect, proofing, formatting boundaries) — a
// straddling run boundary only breaks pattern MATCHING (e.g. finding a
// literal "{{key}}" span), not plain concatenation.
func (s *OpenSlide) Text() (string, error) {
	raw, err := s.tmpl.rawSlideBytes(s.path)
	if err != nil {
		return "", err
	}
	return extractText(raw)
}

// extractText walks raw slide XML collecting every a:t element's chardata
// in document order, concatenating consecutive runs within one a:p with no
// separator (they are one logical line, however PowerPoint split them
// across runs) and separating paragraphs that contain any text with "\n".
func extractText(raw []byte) (string, error) {
	dec := xml.NewDecoder(bytes.NewReader(raw))
	var sb strings.Builder
	inT := false
	paragraphHasText := false

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", errors.Wrap(err, "pptx.extractText")
		}

		switch el := tok.(type) {
		case xml.StartElement:
			switch el.Name.Local {
			case "t":
				inT = true
			case "p":
				if paragraphHasText {
					sb.WriteByte('\n')
				}
				paragraphHasText = false
			}
		case xml.EndElement:
			if el.Name.Local == "t" {
				inT = false
			}
		case xml.CharData:
			if inT {
				sb.Write(el)
				paragraphHasText = true
			}
		}
	}

	return sb.String(), nil
}
