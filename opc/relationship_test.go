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

import "testing"

func TestRelationshipManager_AddGeneratesSequentialIDs(t *testing.T) {
	rm := NewRelationshipManager()

	id1, err := rm.AddImage("media/image1.png")
	if err != nil {
		t.Fatalf("AddImage: %v", err)
	}
	id2, err := rm.AddImage("media/image2.png")
	if err != nil {
		t.Fatalf("AddImage: %v", err)
	}
	if id1 == id2 {
		t.Fatalf("expected distinct IDs, got %s twice", id1)
	}
	if rm.Count() != 2 {
		t.Fatalf("expected 2 relationships, got %d", rm.Count())
	}
}

func TestRelationshipManager_AddImageReusesExistingImageRelToSameTarget(t *testing.T) {
	rm := NewRelationshipManager()

	id1, err := rm.AddImage("../media/image1.png")
	if err != nil {
		t.Fatalf("AddImage: %v", err)
	}
	id2, err := rm.AddImage("../media/image1.png")
	if err != nil {
		t.Fatalf("AddImage: %v", err)
	}
	if id1 != id2 {
		t.Errorf("expected AddImage to reuse the existing rel (same ID), got %s and %s", id1, id2)
	}
	if rm.Count() != 1 {
		t.Errorf("expected exactly 1 relationship after reusing the target, got %d", rm.Count())
	}
}

func TestRelationshipManager_AddImageIgnoresNonImageRelWithSameTarget(t *testing.T) {
	// A same-target relationship of a different type must not shadow an
	// AddImage lookup — the type-scoped search (not a bare GetByTarget)
	// is what guarantees this regardless of map iteration order.
	rm := NewRelationshipManager()

	hlinkID, err := rm.Add(RelTypeHyperlink, "../media/image1.png", "External")
	if err != nil {
		t.Fatalf("Add: %v", err)
	}

	imgID, err := rm.AddImage("../media/image1.png")
	if err != nil {
		t.Fatalf("AddImage: %v", err)
	}
	if imgID == hlinkID {
		t.Errorf("expected AddImage to create its own image relationship, not reuse the hyperlink's ID %s", hlinkID)
	}
	if rm.Count() != 2 {
		t.Errorf("expected 2 distinct relationships (hyperlink + image), got %d", rm.Count())
	}

	rel, err := rm.Get(imgID)
	if err != nil {
		t.Fatalf("Get(%s): %v", imgID, err)
	}
	if rel.Type != RelTypeImage {
		t.Errorf("expected the AddImage relationship to have RelTypeImage, got %s", rel.Type)
	}
}

func TestRelationshipManager_InternalTargetModeOmitted(t *testing.T) {
	rm := NewRelationshipManager()
	id, err := rm.AddImage("media/image1.png")
	if err != nil {
		t.Fatalf("AddImage: %v", err)
	}
	rel, err := rm.Get(id)
	if err != nil {
		t.Fatalf("Get: %v", err)
	}
	if rel.TargetMode != "" {
		t.Errorf("expected Internal relationship to have empty TargetMode, got %q", rel.TargetMode)
	}
}

func TestRelationshipManager_ExternalTargetModeCaseCanonicalized(t *testing.T) {
	// The OPC schema's TargetMode enum is case-sensitive ("External", not
	// "external" or "EXTERNAL"); a caller passing a differently-cased
	// variant must still get the one value Office accepts.
	for _, mode := range []string{"external", "EXTERNAL", "External", "eXternal"} {
		rm := NewRelationshipManager()
		id, err := rm.Add("http://example.com/rel/thing", "target", mode)
		if err != nil {
			t.Fatalf("Add(%q): %v", mode, err)
		}
		rel, err := rm.Get(id)
		if err != nil {
			t.Fatalf("Get: %v", err)
		}
		if rel.TargetMode != "External" {
			t.Errorf("Add(mode=%q): TargetMode = %q, want \"External\"", mode, rel.TargetMode)
		}
	}
}

func TestRelIDLess_TiedNumericValueFallsBackToStringOrder(t *testing.T) {
	// "rId1" and "rId01" parse to the same number; a comparator that treats
	// them as equal would leave sort.Slice (an unstable sort) free to order
	// them differently across calls. The tie must resolve deterministically.
	if !relIDLess("rId01", "rId1") {
		t.Error(`relIDLess("rId01", "rId1") = false, want true (string tiebreak)`)
	}
	if relIDLess("rId1", "rId01") {
		t.Error(`relIDLess("rId1", "rId01") = true, want false (string tiebreak)`)
	}
}

func TestRelationshipManager_ExternalTargetModePreserved(t *testing.T) {
	rm := NewRelationshipManager()
	id, err := rm.AddHyperlink("https://example.com")
	if err != nil {
		t.Fatalf("AddHyperlink: %v", err)
	}
	rel, err := rm.Get(id)
	if err != nil {
		t.Fatalf("Get: %v", err)
	}
	if rel.TargetMode != "External" {
		t.Errorf("expected External TargetMode, got %q", rel.TargetMode)
	}
}

func TestRelationshipManager_AddRejectsEmptyFields(t *testing.T) {
	rm := NewRelationshipManager()
	if _, err := rm.Add("", "target", ""); err == nil {
		t.Error("expected error for empty relType")
	}
	if _, err := rm.Add("type", "", ""); err == nil {
		t.Error("expected error for empty target")
	}
}

func TestRelationshipManager_RegisterExistingAdvancesIDCounter(t *testing.T) {
	rm := NewRelationshipManager()
	if err := rm.RegisterExisting("rId5", "type", "target", ""); err != nil {
		t.Fatalf("RegisterExisting: %v", err)
	}

	next, err := rm.AddImage("media/imageX.png")
	if err != nil {
		t.Fatalf("AddImage: %v", err)
	}
	if next == "rId5" || next == "rId1" {
		t.Errorf("expected a fresh ID past rId5, got %s", next)
	}
}

func TestIDGenerator_SeparatePrefixesDoNotCollide(t *testing.T) {
	g := NewIDGenerator()
	if got, want := g.NextID("image"), "image1"; got != want {
		t.Errorf("NextID(image) = %s, want %s", got, want)
	}
	if got, want := g.NextID("shape"), "shape1"; got != want {
		t.Errorf("NextID(shape) = %s, want %s", got, want)
	}
	if got, want := g.NextID("image"), "image2"; got != want {
		t.Errorf("NextID(image) = %s, want %s", got, want)
	}
}
