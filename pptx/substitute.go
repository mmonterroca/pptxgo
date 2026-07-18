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

// This file implements text substitution over an opened slide's raw XML —
// the one place Template/OpenSlide (open.go) actually mutate content. It
// never parses a slide into the write-only content-model structs (xml.go);
// it byte-splices the original bytes instead, so everything the substitution
// engine doesn't touch — geometry, fills, unrelated paragraphs, whole other
// parts — passes through exactly as loaded.
package pptx

import (
	"bytes"
	"encoding/xml"
	"io"
	"sort"

	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// runSpan is one <a:r> found scanning a slide's raw XML.
type runSpan struct {
	start, end int64 // byte range of the whole <a:r>...</a:r> in the original raw bytes

	// rPrRaw is the exact source bytes of this run's <a:rPr>, or "" if the
	// run has none. Runs are only ever grouped for merging when this
	// matches BYTE FOR BYTE — PowerPoint's own split-run behavior
	// (autocorrect, proofing, formatting boundaries) clones rPr byte-for-
	// byte when it fragments a logical string across runs, so exact
	// equality is the correct test for healing that, not a false
	// restriction. A run with no rPr at all is its own equality class (an
	// unformatted run never merges with a differently- or un-formatted
	// one by coincidence).
	rPrRaw string

	hasT         bool   // whether this run has an <a:t> child at all (schema requires it; tolerated as absent defensively rather than erroring on a malformed run)
	tStart, tEnd int64  // byte range of the FULL <a:t>...</a:t> (or self-closed <a:t/>) span, tags included — see substituteSlideText for why the edit unit is the whole element, not just its inner content
	text         string // decoded (already XML-unescaped) text content of a:t

	// groupBoundaryBefore is true when this run must NOT merge with the
	// immediately preceding run: either a new paragraph started, or a
	// non-run sibling (a:br, a:fld, ...) came between them. CT_RegularTextRun
	// (a:r) only ever contains rPr?+t, so there is no "is this run itself
	// text-only" question the way a naive per-run check might ask —
	// boundaries are entirely about what sits BETWEEN runs at the
	// paragraph level.
	groupBoundaryBefore bool
}

// scanRuns walks raw (a slide's XML) and returns every <a:r> that is a
// direct child of some <a:p>, in document order.
//
// Unlike a naive substring/regex scan for "<a:t", walking with
// xml.Decoder.RawToken() parses complete element boundaries, so "a:t" can
// never be confused with "a:tbl"/"a:tc"/"a:tcPr"/etc. sharing the same
// prefix — RawToken preserves the original element's prefix+local name
// (Name.Space holds the raw prefix string, e.g. "a", not a resolved
// namespace URI — see RawToken's own doc comment) without attempting the
// resolution that would otherwise be needed to compare against a real
// namespace, which this function deliberately never needs.
func scanRuns(raw []byte) ([]*runSpan, error) {
	dec := xml.NewDecoder(bytes.NewReader(raw))

	type frame struct {
		local string
		start int64
	}
	var stack []frame

	var runs []*runSpan
	var cur *runSpan
	var curText bytes.Buffer
	inT := false
	pendingBoundary := true

	for {
		startOffset := dec.InputOffset()
		tok, err := dec.RawToken()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, errors.Wrap(err, "pptx.scanRuns")
		}
		endOffset := dec.InputOffset()

		switch el := tok.(type) {
		case xml.StartElement:
			parent := ""
			if len(stack) > 0 {
				parent = stack[len(stack)-1].local
			}
			switch {
			case el.Name.Local == "p":
				pendingBoundary = true
			case el.Name.Local == "r" && parent == "p" && cur == nil:
				cur = &runSpan{start: startOffset, groupBoundaryBefore: pendingBoundary}
				pendingBoundary = false
			case cur != nil && el.Name.Local == "t":
				cur.hasT = true
				cur.tStart = startOffset
				inT = true
				curText.Reset()
			case parent == "p" && el.Name.Local != "r":
				// A non-run paragraph child (a:br, a:fld, a:pPr, a:endParaRPr, ...)
				// breaks consolidation between the runs before and after it.
				pendingBoundary = true
			}
			stack = append(stack, frame{local: el.Name.Local, start: startOffset})

		case xml.EndElement:
			top := stack[len(stack)-1]
			stack = stack[:len(stack)-1]

			switch {
			case cur != nil && top.local == "rPr":
				cur.rPrRaw = string(raw[top.start:endOffset])
			case cur != nil && top.local == "t":
				cur.tEnd = endOffset
				cur.text = curText.String()
				inT = false
			case cur != nil && top.local == "r":
				cur.end = endOffset
				runs = append(runs, cur)
				cur = nil
			}

		case xml.CharData:
			if inT {
				curText.Write(el)
			}
		}
	}

	return runs, nil
}

// runGroup is a maximal run of consecutive runSpans (within one paragraph,
// no intervening non-run sibling) sharing byte-identical rPr — the unit
// substitution actually operates on, so a pattern split across runs (e.g.
// "{{na" + "me}}") is matched against the CONCATENATED text, not each
// run's own fragment.
type runGroup struct {
	text    string
	members []*runSpan // members[0] receives any substituted text; the rest are blanked if the group's text changes — see substituteSlideText
}

// groupRuns partitions runs (in document order, as returned by scanRuns)
// into runGroups. A run with no <a:t> at all contributes nothing and is
// skipped entirely — there is no text to read or write, so it is left
// untouched by construction (substituteSlideText never generates an edit
// for it).
func groupRuns(runs []*runSpan) []*runGroup {
	var groups []*runGroup
	for _, r := range runs {
		if !r.hasT {
			continue
		}
		var g *runGroup
		if n := len(groups); n > 0 {
			last := groups[n-1]
			lastMember := last.members[len(last.members)-1]
			if !r.groupBoundaryBefore && lastMember.rPrRaw == r.rPrRaw {
				g = last
			}
		}
		if g == nil {
			g = &runGroup{}
			groups = append(groups, g)
		}
		g.text += r.text
		g.members = append(g.members, r)
	}
	return groups
}

// substituteSlideText scans raw, applies transform to each run-group's
// concatenated text, and — only for groups where transform actually
// changed the text — splices the result back into a fresh copy of raw.
// Groups transform leaves unchanged are never touched at all, even if
// they span multiple runs: a pure inspection pass (Text(), PlaceholderNames)
// or a Merge/Replace call that matches nothing must never rewrite a
// slide's run structure as a side effect.
//
// Returns the new bytes (== raw, unmodified, if nothing changed) and how
// many groups were actually changed.
func substituteSlideText(raw []byte, transform func(string) string) ([]byte, int, error) {
	runs, err := scanRuns(raw)
	if err != nil {
		return nil, 0, err
	}
	groups := groupRuns(runs)

	type edit struct {
		start, end  int64
		replacement []byte
	}
	var edits []edit
	changed := 0

	for _, g := range groups {
		newText := transform(g.text)
		if newText == g.text {
			continue
		}
		changed++

		first := g.members[0]
		edits = append(edits, edit{start: first.tStart, end: first.tEnd, replacement: renderAT(newText)})
		for _, m := range g.members[1:] {
			edits = append(edits, edit{start: m.tStart, end: m.tEnd, replacement: renderAT("")})
		}
	}

	if len(edits) == 0 {
		return raw, 0, nil
	}

	sort.Slice(edits, func(i, j int) bool { return edits[i].start < edits[j].start })

	var out bytes.Buffer
	cursor := int64(0)
	for _, e := range edits {
		out.Write(raw[cursor:e.start])
		out.Write(e.replacement)
		cursor = e.end
	}
	out.Write(raw[cursor:])

	return out.Bytes(), changed, nil
}

// renderAT renders a fresh <a:t>...</a:t> element for text: XML-escaped,
// and always carrying xml:space="preserve" regardless of whether text
// itself has leading/trailing whitespace. Always adding it (rather than
// only when text's edges need it) sidesteps a fragile per-case whitespace
// check for one extra, harmless attribute — PowerPoint already preserves
// a:t's internal whitespace verbatim as a DrawingML text-layout
// convention, so a superfluous xml:space="preserve" changes nothing a
// caller would observe.
func renderAT(text string) []byte {
	var buf bytes.Buffer
	buf.WriteString(`<a:t xml:space="preserve">`)
	xml.EscapeText(&buf, []byte(text))
	buf.WriteString(`</a:t>`)
	return buf.Bytes()
}
