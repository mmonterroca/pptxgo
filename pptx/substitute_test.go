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
)

// splitPlaceholderSlideXML is a HAND-CRAFTED slide fragment (not
// pptxgo-generated, not from any external fixture file) simulating the
// exact real-PowerPoint failure mode this substitution engine exists to
// heal: "{{client_name}}" fragmented across two <a:r> runs with
// byte-identical <a:rPr> (autocorrect/proofing splits a logical string
// this way — see runSpan.rPrRaw's own doc comment), followed by a third
// run with genuinely DIFFERENT formatting (bold) that must stay separate.
const splitPlaceholderSlideXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:p>
            <a:r><a:rPr lang="en-US" dirty="0"/><a:t>{{client_</a:t></a:r>
            <a:r><a:rPr lang="en-US" dirty="0"/><a:t>name}}</a:t></a:r>
            <a:r><a:rPr lang="en-US" b="1" dirty="0"/><a:t> Quarterly</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`

func TestSubstituteSlideText_HealsPlaceholderSplitAcrossRuns(t *testing.T) {
	out, changed, err := substituteSlideText([]byte(splitPlaceholderSlideXML), func(text string) string {
		return strings.ReplaceAll(text, "{{client_name}}", "Acme Corp")
	})
	if err != nil {
		t.Fatalf("substituteSlideText: %v", err)
	}
	if changed != 1 {
		t.Errorf("expected exactly 1 changed group (the merged placeholder run), got %d", changed)
	}

	got := string(out)
	if !strings.Contains(got, "Acme Corp") {
		t.Fatalf("expected the healed substitution to appear, got %s", got)
	}
	if strings.Contains(got, "{{client_") || strings.Contains(got, "name}}") {
		t.Errorf("expected no leftover placeholder fragments after healing a cross-run split, got %s", got)
	}
	// The un-merged, differently-formatted third run's own text must
	// survive untouched.
	if !strings.Contains(got, " Quarterly") {
		t.Errorf("expected the bold run's own text to survive untouched, got %s", got)
	}
	if !strings.Contains(got, `b="1"`) {
		t.Errorf("expected the bold run's own rPr to survive untouched, got %s", got)
	}

	// The output must still be well-formed XML, and extracting its text
	// must read back the fully healed, concatenated result.
	var v any
	if err := xml.Unmarshal(out, &v); err != nil {
		t.Fatalf("output is not well-formed XML: %v", err)
	}
	text, err := extractText(out)
	if err != nil {
		t.Fatalf("extractText(out): %v", err)
	}
	if text != "Acme Corp Quarterly" {
		t.Errorf("extractText(out) = %q, want %q", text, "Acme Corp Quarterly")
	}
}

func TestSubstituteSlideText_DoesNotMergeAcrossDifferentRPr(t *testing.T) {
	// Two adjacent runs whose text WOULD form a placeholder if
	// concatenated, but whose formatting genuinely differs, must NOT be
	// treated as one match -- over-merging would silently destroy the
	// second run's distinct formatting (see runGroup's own doc comment).
	xmlStr := `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p>` +
		`<a:r><a:rPr lang="en-US"/><a:t>{{cli</a:t></a:r>` +
		`<a:r><a:rPr lang="en-US" b="1"/><a:t>ent_name}}</a:t></a:r>` +
		`</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`

	out, changed, err := substituteSlideText([]byte(xmlStr), func(text string) string {
		return strings.ReplaceAll(text, "{{client_name}}", "Acme Corp")
	})
	if err != nil {
		t.Fatalf("substituteSlideText: %v", err)
	}
	if changed != 0 {
		t.Errorf("expected no group to match (formatting differs, so the two runs never concatenate into one group), got %d changed", changed)
	}
	if string(out) != xmlStr {
		t.Errorf("expected byte-identical output when nothing matches, got %s", out)
	}
}

func TestSubstituteSlideText_ABreakInterruptsConsolidation(t *testing.T) {
	// a:br sits between two identically-formatted runs -- they must not
	// merge across it, since concatenating across a line break would
	// silently glue text from two different lines together.
	xmlStr := `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p>` +
		`<a:r><a:rPr lang="en-US"/><a:t>{{cli</a:t></a:r>` +
		`<a:br/>` +
		`<a:r><a:rPr lang="en-US"/><a:t>ent_name}}</a:t></a:r>` +
		`</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`

	_, changed, err := substituteSlideText([]byte(xmlStr), func(text string) string {
		return strings.ReplaceAll(text, "{{client_name}}", "Acme Corp")
	})
	if err != nil {
		t.Fatalf("substituteSlideText: %v", err)
	}
	if changed != 0 {
		t.Errorf("expected a:br to break consolidation (no cross-break match), got %d changed", changed)
	}
}

func TestSubstituteSlideText_ANewParagraphInterruptsConsolidation(t *testing.T) {
	xmlStr := `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody>` +
		`<a:p><a:r><a:rPr lang="en-US"/><a:t>{{cli</a:t></a:r></a:p>` +
		`<a:p><a:r><a:rPr lang="en-US"/><a:t>ent_name}}</a:t></a:r></a:p>` +
		`</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`

	_, changed, err := substituteSlideText([]byte(xmlStr), func(text string) string {
		return strings.ReplaceAll(text, "{{client_name}}", "Acme Corp")
	})
	if err != nil {
		t.Fatalf("substituteSlideText: %v", err)
	}
	if changed != 0 {
		t.Errorf("expected a paragraph boundary to break consolidation (no cross-paragraph match), got %d changed", changed)
	}
}

func TestSubstituteSlideText_NoMatchLeavesBytesIdentical(t *testing.T) {
	out, changed, err := substituteSlideText([]byte(splitPlaceholderSlideXML), func(text string) string {
		return text // identity transform -- simulates Text()/PlaceholderNames-style inspection with no actual edit
	})
	if err != nil {
		t.Fatalf("substituteSlideText: %v", err)
	}
	if changed != 0 {
		t.Errorf("expected 0 changed groups for an identity transform, got %d", changed)
	}
	if string(out) != splitPlaceholderSlideXML {
		t.Error("expected byte-identical output when the transform changes nothing, even though a mergeable multi-run group exists")
	}
}

func TestSubstituteSlideText_DistinguishesATFromATblAndATc(t *testing.T) {
	// Regression for the exact ambiguity a naive "<a:t" substring/regex
	// scan would risk: a:tbl/a:tc/a:tcPr/a:tblPr/a:tblGrid all share the
	// "a:t" prefix. A real tokenizer (RawToken) parses element boundaries,
	// so this must never confuse a table for a text run.
	xmlStr := `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:graphicFrame><a:graphic><a:graphicData>` +
		`<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="100"/></a:tblGrid><a:tr h="50"><a:tc><a:txBody><a:p>` +
		`<a:r><a:rPr lang="en-US"/><a:t>{{client_name}}</a:t></a:r>` +
		`</a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl>` +
		`</a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>`

	out, changed, err := substituteSlideText([]byte(xmlStr), func(text string) string {
		return strings.ReplaceAll(text, "{{client_name}}", "Acme Corp")
	})
	if err != nil {
		t.Fatalf("substituteSlideText: %v", err)
	}
	if changed != 1 {
		t.Fatalf("expected exactly 1 changed group (the table cell's own run), got %d", changed)
	}
	got := string(out)
	if !strings.Contains(got, "Acme Corp") {
		t.Errorf("expected the table cell's placeholder to be substituted, got %s", got)
	}
	if !strings.Contains(got, "<a:tblGrid>") || !strings.Contains(got, "<a:tcPr/>") {
		t.Errorf("expected table structure to survive untouched, got %s", got)
	}

	var v any
	if err := xml.Unmarshal(out, &v); err != nil {
		t.Fatalf("output is not well-formed XML: %v", err)
	}
}

func TestScanRuns_CapturesSpansForTheHandCraftedFixture(t *testing.T) {
	runs, err := scanRuns([]byte(splitPlaceholderSlideXML))
	if err != nil {
		t.Fatalf("scanRuns: %v", err)
	}
	if len(runs) != 3 {
		t.Fatalf("expected 3 runs, got %d", len(runs))
	}
	if runs[0].text != "{{client_" || runs[1].text != "name}}" || runs[2].text != " Quarterly" {
		t.Errorf("unexpected run texts: %q, %q, %q", runs[0].text, runs[1].text, runs[2].text)
	}
	if runs[0].rPrRaw != runs[1].rPrRaw {
		t.Errorf("expected runs[0] and runs[1] to share byte-identical rPr, got %q vs %q", runs[0].rPrRaw, runs[1].rPrRaw)
	}
	if runs[1].rPrRaw == runs[2].rPrRaw {
		t.Errorf("expected runs[2]'s bold rPr to differ from runs[1]'s, got both %q", runs[1].rPrRaw)
	}
	if runs[0].groupBoundaryBefore != true {
		t.Errorf("expected the first run in a paragraph to have groupBoundaryBefore=true, got false")
	}
	if runs[1].groupBoundaryBefore {
		t.Errorf("expected the second run (no intervening sibling) to have groupBoundaryBefore=false, got true")
	}
}

func TestGroupRuns_MergesOnlyByteIdenticalRPrWithinOneParagraph(t *testing.T) {
	runs, err := scanRuns([]byte(splitPlaceholderSlideXML))
	if err != nil {
		t.Fatalf("scanRuns: %v", err)
	}
	groups := groupRuns(runs)
	if len(groups) != 2 {
		t.Fatalf("expected 2 groups (runs 0+1 merged, run 2 separate), got %d", len(groups))
	}
	if groups[0].text != "{{client_name}}" {
		t.Errorf("groups[0].text = %q, want %q", groups[0].text, "{{client_name}}")
	}
	if len(groups[0].members) != 2 {
		t.Errorf("expected groups[0] to have 2 members, got %d", len(groups[0].members))
	}
	if groups[1].text != " Quarterly" {
		t.Errorf("groups[1].text = %q, want %q", groups[1].text, " Quarterly")
	}
	if len(groups[1].members) != 1 {
		t.Errorf("expected groups[1] to have 1 member, got %d", len(groups[1].members))
	}
}
