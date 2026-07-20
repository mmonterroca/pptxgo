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

package drawingml

import (
	"strings"
	"testing"
)

func TestGroupXfrm_ChildOrderIsOffExtChOffChExt(t *testing.T) {
	x := &GroupXfrm{
		Off:   &Off{X: 100, Y: 200},
		Ext:   &Ext{Cx: 300, Cy: 400},
		ChOff: &ChOff{X: 100, Y: 200},
		ChExt: &ChExt{Cx: 300, Cy: 400},
	}
	got := marshal(t, x)

	offIdx := strings.Index(got, "<a:off")
	extIdx := strings.Index(got, "<a:ext")
	chOffIdx := strings.Index(got, "<a:chOff")
	chExtIdx := strings.Index(got, "<a:chExt")
	if offIdx == -1 || extIdx == -1 || chOffIdx == -1 || chExtIdx == -1 {
		t.Fatalf("expected off, ext, chOff, chExt all present, got %s", got)
	}
	if !(offIdx < extIdx && extIdx < chOffIdx && chOffIdx < chExtIdx) {
		t.Errorf("expected off < ext < chOff < chExt (CT_GroupTransform2D sequence), got %s", got)
	}
	if !strings.Contains(got, `<a:chOff x="100" y="200">`) {
		t.Errorf("expected chOff attrs, got %s", got)
	}
	if !strings.Contains(got, `<a:chExt cx="300" cy="400">`) {
		t.Errorf("expected chExt attrs, got %s", got)
	}
}

func TestGroupXfrm_SharesATagWithXfrmButIsADistinctType(t *testing.T) {
	// GroupXfrm and Xfrm both marshal as "a:xfrm" (used in different parent
	// contexts, p:grpSpPr vs p:spPr) — confirm GroupXfrm's own chOff/chExt
	// still marshal correctly under that shared tag name, i.e. the shared
	// tag alone doesn't confuse encoding/xml into using Xfrm's shape.
	got := marshal(t, &GroupXfrm{ChOff: &ChOff{X: 1, Y: 2}, ChExt: &ChExt{Cx: 3, Cy: 4}})
	if !strings.HasPrefix(got, "<a:xfrm>") {
		t.Errorf("expected GroupXfrm to marshal as a:xfrm, got %s", got)
	}
	if !strings.Contains(got, "<a:chOff") {
		t.Errorf("expected chOff present, got %s", got)
	}
}

func TestChOffChExt_MarshalUnderOwnDistinctTags(t *testing.T) {
	// Regression: ChOff/ChExt must NOT reuse Off/Ext directly — Off/Ext's
	// own fixed "a:off"/"a:ext" XMLName would win over any field tag that
	// tried to rename them to "a:chOff"/"a:chExt" (encoding/xml errors on
	// that mismatch, the same class of conflict TcBorderLn/LineEnd already
	// document).
	got := marshal(t, &ChOff{X: 5, Y: 6})
	if !strings.Contains(got, `<a:chOff x="5" y="6">`) {
		t.Errorf("expected a:chOff, got %s", got)
	}
	got = marshal(t, &ChExt{Cx: 7, Cy: 8})
	if !strings.Contains(got, `<a:chExt cx="7" cy="8">`) {
		t.Errorf("expected a:chExt, got %s", got)
	}
}

func TestStCxnEndCxn_MarshalUnderOwnDistinctTagsWithIDAndIdx(t *testing.T) {
	got := marshal(t, &StCxn{ID: 2, Idx: 3})
	if !strings.Contains(got, `<a:stCxn id="2" idx="3">`) {
		t.Errorf("expected a:stCxn with id and idx, got %s", got)
	}
	got = marshal(t, &EndCxn{ID: 5, Idx: 1})
	if !strings.Contains(got, `<a:endCxn id="5" idx="1">`) {
		t.Errorf("expected a:endCxn with id and idx, got %s", got)
	}
}
