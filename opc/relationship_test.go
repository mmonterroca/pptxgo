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
