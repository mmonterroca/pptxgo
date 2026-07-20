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
	"encoding/xml"
	"strings"
	"testing"
	"time"
)

func TestWithMetadataOptions_PopulateCoreAndAppProps(t *testing.T) {
	p := New(
		WithTitle("Quarterly Review"),
		WithAuthor("Ada Lovelace"),
		WithSubject("Q3 results"),
		WithKeywords("finance, quarterly"),
		WithDescription("Internal review deck"),
		WithCompany("Acme Corp"),
	)
	p.AddSlide()
	files := generateFrom(t, p)

	core := string(files["docProps/core.xml"])
	for _, want := range []string{
		"<dc:title>Quarterly Review</dc:title>",
		"<dc:creator>Ada Lovelace</dc:creator>",
		"<dc:subject>Q3 results</dc:subject>",
		"<cp:keywords>finance, quarterly</cp:keywords>",
		"<dc:description>Internal review deck</dc:description>",
		// LastModifiedBy defaults to the author.
		"<cp:lastModifiedBy>Ada Lovelace</cp:lastModifiedBy>",
	} {
		if !strings.Contains(core, want) {
			t.Errorf("expected core.xml to contain %q, got:\n%s", want, core)
		}
	}

	app := string(files["docProps/app.xml"])
	if !strings.Contains(app, "<Company>Acme Corp</Company>") {
		t.Errorf("expected app.xml Company, got:\n%s", app)
	}
	// CT_Properties is a strict sequence: Company must precede Application.
	companyIdx := strings.Index(app, "<Company>")
	appIdx := strings.Index(app, "<Application>")
	if companyIdx == -1 || appIdx == -1 || companyIdx > appIdx {
		t.Errorf("expected Company before Application (schema sequence order), got:\n%s", app)
	}
}

func TestWithMetadata_TimestampsEmitW3CDTF(t *testing.T) {
	created := time.Date(2026, 7, 19, 9, 30, 0, 0, time.UTC)
	modified := time.Date(2026, 7, 20, 14, 0, 0, 0, time.UTC)
	p := New(WithMetadata(Metadata{
		Title:    "Timed",
		Created:  created,
		Modified: modified,
	}))
	p.AddSlide()
	core := string(generateFrom(t, p)["docProps/core.xml"])

	for _, want := range []string{
		`<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-19T09:30:00Z</dcterms:created>`,
		`<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-20T14:00:00Z</dcterms:modified>`,
	} {
		if !strings.Contains(core, want) {
			t.Errorf("expected core.xml to contain %q, got:\n%s", want, core)
		}
	}
}

func TestWithMetadata_ExplicitLastModifiedByWins(t *testing.T) {
	p := New(WithMetadata(Metadata{Creator: "Author", LastModifiedBy: "Editor"}))
	p.AddSlide()
	core := string(generateFrom(t, p)["docProps/core.xml"])

	if !strings.Contains(core, "<cp:lastModifiedBy>Editor</cp:lastModifiedBy>") {
		t.Errorf("expected explicit LastModifiedBy to win over the author default, got:\n%s", core)
	}
}

func TestSingleFieldOptionsComposeWithWithMetadata(t *testing.T) {
	// A later single-field option overrides the field set by WithMetadata.
	p := New(
		WithMetadata(Metadata{Title: "Original", Creator: "Author"}),
		WithTitle("Overridden"),
	)
	p.AddSlide()
	core := string(generateFrom(t, p)["docProps/core.xml"])

	if !strings.Contains(core, "<dc:title>Overridden</dc:title>") {
		t.Errorf("expected WithTitle to override WithMetadata's title, got:\n%s", core)
	}
	if !strings.Contains(core, "<dc:creator>Author</dc:creator>") {
		t.Errorf("expected WithMetadata's creator to survive, got:\n%s", core)
	}
}

func TestNoMetadata_ProducesWellFormedMinimalProps(t *testing.T) {
	p := New()
	p.AddSlide()
	files := generateFrom(t, p)

	core := files["docProps/core.xml"]
	app := files["docProps/app.xml"]

	var probe any
	if err := xml.Unmarshal(core, &probe); err != nil {
		t.Errorf("core.xml not well-formed: %v", err)
	}
	if err := xml.Unmarshal(app, &probe); err != nil {
		t.Errorf("app.xml not well-formed: %v", err)
	}
	if strings.Contains(string(core), "<dc:title>") {
		t.Errorf("expected no dc:title with no metadata, got:\n%s", core)
	}
	if strings.Contains(string(app), "<Company>") {
		t.Errorf("expected no Company with no metadata, got:\n%s", app)
	}
	if !strings.Contains(string(app), "<Application>pptxgo</Application>") {
		t.Errorf("expected app.xml to still identify pptxgo, got:\n%s", app)
	}
}
