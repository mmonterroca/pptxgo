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
	"path"
	"sort"
	"strings"
	"sync"
)

// Part is a single piece of an OPC package: a path inside the ZIP container,
// its declared content type, and its content.
//
// Content comes from exactly one of two places, and the package does not
// care which: Raw holds verbatim bytes (typically read from an existing
// file being preserved through a template workflow), and Value holds a Go
// value to be marshaled with encoding/xml at write time (typically a part
// being generated from scratch). Raw takes precedence when both are set.
// This is what makes "generate from scratch" and "open a template and
// replace some parts" the same code path: a Package never distinguishes
// between a part it built and a part it loaded.
type Part struct {
	Path        string
	ContentType string
	Raw         []byte
	Value       any

	// noOverride is true for parts whose content type is fully described by
	// an extension-level Default in [Content_Types].xml (e.g. media files),
	// so no per-part Override is required.
	noOverride bool
}

// Package is an in-memory, format-agnostic OPC package: a set of parts plus
// the relationships between them (and from the package root). It knows
// nothing about what any part means — building a valid PresentationML,
// WordprocessingML, or SpreadsheetML document out of it is entirely the
// job of the caller.
type Package struct {
	mu       sync.Mutex
	parts    map[string]*Part
	order    []string                        // insertion order, for deterministic output
	rels     map[string]*RelationshipManager // key: owner part path, "" = root
	defaults map[string]string               // extension (no dot) -> content type
	ids      *IDGenerator
}

// NewPackage creates an empty package. [Content_Types].xml Defaults for the
// "rels" and "xml" extensions are seeded automatically, matching every real
// OOXML package.
func NewPackage() *Package {
	return &Package{
		parts: make(map[string]*Part),
		rels:  make(map[string]*RelationshipManager),
		defaults: map[string]string{
			"rels": ContentTypeRelationships,
			"xml":  ContentTypeXML,
		},
		ids: NewIDGenerator(),
	}
}

// IDs returns the package's shared ID generator, for identifiers that must
// be unique across the whole package — media file names, shape IDs, and the
// like. Relationship IDs are NOT minted here: they are scoped per owning
// part (see Relationships) and each RelationshipManager numbers its own,
// so mixing this generator into an rId would break that per-part scoping.
func (p *Package) IDs() *IDGenerator {
	return p.ids
}

// normalizePartPath converts a caller-supplied path into the slash-separated
// form OPC part paths require inside the ZIP container, regardless of the
// host OS: backslashes (as filepath.Join produces on Windows) become
// forward slashes, and any leading slash is trimmed.
func normalizePartPath(pth string) string {
	pth = strings.ReplaceAll(pth, `\`, "/")
	return strings.TrimPrefix(pth, "/")
}

// AddPart registers a part to be generated from value at write time
// (marshaled with encoding/xml). It receives an Override entry in
// [Content_Types].xml, since its content type is specific to this part
// and not derivable from its extension alone.
func (p *Package) AddPart(pth, contentType string, value any) *Part {
	return p.addPart(&Part{Path: normalizePartPath(pth), ContentType: contentType, Value: value})
}

// AddRawPart registers a part whose bytes are already fully formed — e.g.
// loaded verbatim from an existing package being used as a template, or a
// hand-authored literal such as a default theme. Like AddPart, it receives
// a Content_Types Override.
func (p *Package) AddRawPart(pth, contentType string, raw []byte) *Part {
	return p.addPart(&Part{Path: normalizePartPath(pth), ContentType: contentType, Raw: raw})
}

// AddMediaPart registers a binary media part (an image, typically). When the
// path has a file extension AND no other part has already claimed that
// extension for a different content type, it is declared via a
// Content_Types Default for that extension — exactly as Office represents
// media, so many images of the same type share one entry rather than one
// Override each. Otherwise (no extension, or the extension's Default
// already means something else) the part falls back to a per-part Override.
// Either way its content type is always declared, never silently omitted or
// silently misdeclared — both of which Office rejects as corrupt.
func (p *Package) AddMediaPart(pth, contentType string, data []byte) *Part {
	pth = normalizePartPath(pth)
	ext := strings.TrimPrefix(strings.ToLower(path.Ext(pth)), ".")

	coveredByDefault := false
	if ext != "" {
		p.mu.Lock()
		if existing, exists := p.defaults[ext]; !exists {
			p.defaults[ext] = contentType
			coveredByDefault = true
		} else if existing == contentType {
			coveredByDefault = true
		}
		p.mu.Unlock()
	}

	part := &Part{Path: pth, ContentType: contentType, Raw: data, noOverride: coveredByDefault}
	return p.addPart(part)
}

func (p *Package) addPart(part *Part) *Part {
	p.mu.Lock()
	defer p.mu.Unlock()

	if _, exists := p.parts[part.Path]; !exists {
		p.order = append(p.order, part.Path)
	}
	p.parts[part.Path] = part
	return part
}

// Part returns the part at pth, if any.
func (p *Package) Part(pth string) (*Part, bool) {
	p.mu.Lock()
	defer p.mu.Unlock()
	part, ok := p.parts[normalizePartPath(pth)]
	return part, ok
}

// Parts returns every registered part, in the order they were added.
func (p *Package) Parts() []*Part {
	p.mu.Lock()
	defer p.mu.Unlock()

	out := make([]*Part, 0, len(p.order))
	for _, pth := range p.order {
		out = append(out, p.parts[pth])
	}
	return out
}

// Relationships returns the relationship manager that owns relationships
// originating from ownerPath, creating it on first use. Pass "" for the
// package root (produces _rels/.rels); pass a part's path (e.g.
// "ppt/slides/slide1.xml") for that part's own relationships.
//
// Per the OPC spec, relationship IDs are scoped per owner: rId1 in one
// part's .rels is unrelated to rId1 anywhere else. A single shared manager
// would make that impossible to model correctly, which is exactly the class
// of bug (rIds resolved against the wrong scope) that motivates keeping
// this per-owner from the start rather than retrofitting it later.
func (p *Package) Relationships(ownerPath string) *RelationshipManager {
	ownerPath = normalizePartPath(ownerPath)

	p.mu.Lock()
	defer p.mu.Unlock()

	rm, ok := p.rels[ownerPath]
	if !ok {
		rm = NewRelationshipManager()
		p.rels[ownerPath] = rm
	}
	return rm
}

// hasRelationships reports whether ownerPath has a non-empty relationship
// manager, without creating one as a side effect.
func (p *Package) hasRelationships(ownerPath string) bool {
	ownerPath = normalizePartPath(ownerPath)

	p.mu.Lock()
	rm, ok := p.rels[ownerPath]
	p.mu.Unlock()
	return ok && rm.Count() > 0
}

// relsPathForPart returns the .rels part path that holds the relationships
// owned by partPath, following the OPC convention: a sibling "_rels"
// directory containing "<basename>.rels".
func relsPathForPart(partPath string) string {
	dir := path.Dir(partPath)
	base := path.Base(partPath)
	if dir == "." {
		return "_rels/" + base + ".rels"
	}
	return dir + "/_rels/" + base + ".rels"
}

// contentTypesXML derives [Content_Types].xml from the package's registered
// Defaults and every part that requires an Override. It is always computed
// from the current part set — it is never hand-maintained by the caller —
// so it is impossible for a part to exist without a matching entry, or vice
// versa.
func (p *Package) contentTypesXML() *XMLContentTypes {
	p.mu.Lock()
	defer p.mu.Unlock()

	ct := &XMLContentTypes{Xmlns: NamespaceContentTypes}

	extensions := make([]string, 0, len(p.defaults))
	for ext := range p.defaults {
		extensions = append(extensions, ext)
	}
	sort.Strings(extensions)
	for _, ext := range extensions {
		ct.Defaults = append(ct.Defaults, &XMLDefault{Extension: ext, ContentType: p.defaults[ext]})
	}

	for _, pth := range p.order {
		part := p.parts[pth]
		if part.noOverride {
			continue
		}
		ct.Overrides = append(ct.Overrides, &XMLOverride{PartName: "/" + part.Path, ContentType: part.ContentType})
	}

	return ct
}
