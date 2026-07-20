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
	"strings"
	"testing"
)

func TestNotes_CreatesNotesMasterAndNotesSlideParts(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.Notes("Remember to slow down on this slide.")

	files := generateFrom(t, p)

	for _, want := range []string{
		"ppt/notesMasters/notesMaster1.xml",
		"ppt/notesMasters/_rels/notesMaster1.xml.rels",
		"ppt/notesSlides/notesSlide1.xml",
		"ppt/notesSlides/_rels/notesSlide1.xml.rels",
	} {
		if _, ok := files[want]; !ok {
			t.Errorf("missing expected notes part %s", want)
		}
	}

	notes := string(files["ppt/notesSlides/notesSlide1.xml"])
	if !strings.Contains(notes, "<p:notes") {
		t.Errorf("expected p:notes root, got %s", notes)
	}
	if !strings.Contains(notes, `type="body"`) {
		t.Errorf("expected a body placeholder in the notes slide, got %s", notes)
	}
	if !strings.Contains(notes, "Remember to slow down on this slide.") {
		t.Errorf("expected the notes text in notesSlide1.xml, got %s", notes)
	}

	master := string(files["ppt/notesMasters/notesMaster1.xml"])
	if !strings.Contains(master, "<p:notesMaster") || !strings.Contains(master, "<p:clrMap") {
		t.Errorf("expected a p:notesMaster with a clrMap, got %s", master)
	}
}

func TestNotes_WiresRelationshipsAndContentTypes(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.Notes("A note.")
	files := generateFrom(t, p)

	// Slide -> notes slide.
	slideRels := string(files["ppt/slides/_rels/slide1.xml.rels"])
	if !strings.Contains(slideRels, `Target="../notesSlides/notesSlide1.xml"`) {
		t.Errorf("expected slide -> notesSlide relationship, got %s", slideRels)
	}

	// Notes slide -> notes master and -> slide.
	notesRels := string(files["ppt/notesSlides/_rels/notesSlide1.xml.rels"])
	if !strings.Contains(notesRels, `Target="../notesMasters/notesMaster1.xml"`) {
		t.Errorf("expected notesSlide -> notesMaster relationship, got %s", notesRels)
	}
	if !strings.Contains(notesRels, `Target="../slides/slide1.xml"`) {
		t.Errorf("expected notesSlide -> slide relationship, got %s", notesRels)
	}

	// Notes master -> its own theme part (theme2, not the slide master's theme1).
	if _, ok := files["ppt/theme/theme2.xml"]; !ok {
		t.Error("expected the notes master to get its own ppt/theme/theme2.xml part")
	}
	masterRels := string(files["ppt/notesMasters/_rels/notesMaster1.xml.rels"])
	if !strings.Contains(masterRels, `Target="../theme/theme2.xml"`) {
		t.Errorf("expected notesMaster -> theme2 relationship, got %s", masterRels)
	}

	// Presentation -> notes master.
	presRels := string(files["ppt/_rels/presentation.xml.rels"])
	if !strings.Contains(presRels, `Target="notesMasters/notesMaster1.xml"`) {
		t.Errorf("expected presentation -> notesMaster relationship, got %s", presRels)
	}

	// Content types registered for both parts.
	ct := string(files["[Content_Types].xml"])
	for _, want := range []string{
		`PartName="/ppt/notesMasters/notesMaster1.xml"`,
		`PartName="/ppt/notesSlides/notesSlide1.xml"`,
	} {
		if !strings.Contains(ct, want) {
			t.Errorf("expected [Content_Types].xml override %s, got %s", want, ct)
		}
	}
}

func TestNotes_PresentationListsNotesMasterInSchemaOrder(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.Notes("note")
	pres := string(generateFrom(t, p)["ppt/presentation.xml"])

	smIdx := strings.Index(pres, "<p:sldMasterIdLst>")
	nmIdx := strings.Index(pres, "<p:notesMasterIdLst>")
	slIdx := strings.Index(pres, "<p:sldIdLst>")
	if smIdx == -1 || nmIdx == -1 || slIdx == -1 {
		t.Fatalf("expected sldMasterIdLst, notesMasterIdLst, sldIdLst all present, got %s", pres)
	}
	if !(smIdx < nmIdx && nmIdx < slIdx) {
		t.Errorf("expected order sldMasterIdLst < notesMasterIdLst < sldIdLst, got %s", pres)
	}
}

func TestNotes_NotCalledEmitsNoNotesParts(t *testing.T) {
	p := New()
	p.AddSlide()
	files := generateFrom(t, p)

	for name := range files {
		if strings.Contains(name, "notesMaster") || strings.Contains(name, "notesSlide") {
			t.Errorf("expected no notes parts when Notes is never called, got %s", name)
		}
	}
	if strings.Contains(string(files["ppt/presentation.xml"]), "notesMasterIdLst") {
		t.Error("expected no notesMasterIdLst when Notes is never called")
	}
}

func TestNotes_SharesOneNotesMasterAcrossSlides(t *testing.T) {
	p := New()
	p.AddSlide().Notes("first")
	p.AddSlide().Notes("second")
	files := generateFrom(t, p)

	if _, ok := files["ppt/notesSlides/notesSlide1.xml"]; !ok {
		t.Error("missing notesSlide1.xml")
	}
	if _, ok := files["ppt/notesSlides/notesSlide2.xml"]; !ok {
		t.Error("missing notesSlide2.xml")
	}
	if _, ok := files["ppt/notesMasters/notesMaster2.xml"]; ok {
		t.Error("expected a single shared notes master, but a second one was created")
	}
}

func TestNotes_MultilineSplitsIntoBreaks(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.Notes("line one\nline two")
	notes := string(generateFrom(t, p)["ppt/notesSlides/notesSlide1.xml"])

	if !strings.Contains(notes, "<a:br>") && !strings.Contains(notes, "<a:br/>") {
		t.Errorf("expected an a:br between note lines, got %s", notes)
	}
	if !strings.Contains(notes, "line one") || !strings.Contains(notes, "line two") {
		t.Errorf("expected both note lines, got %s", notes)
	}
}

func TestNotes_RepeatCallAppendsParagraph(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.Notes("first paragraph")
	s.Notes("second paragraph")
	notes := string(generateFrom(t, p)["ppt/notesSlides/notesSlide1.xml"])

	if !strings.Contains(notes, "first paragraph") || !strings.Contains(notes, "second paragraph") {
		t.Errorf("expected both note paragraphs, got %s", notes)
	}
	// Two <a:p> paragraphs in the notes body.
	if strings.Count(notes, "<a:p>") < 2 {
		t.Errorf("expected at least 2 paragraphs after a repeat Notes call, got %s", notes)
	}
	// Still a single notes slide part for the slide.
	files := generateFrom(t, p)
	if _, ok := files["ppt/notesSlides/notesSlide2.xml"]; ok {
		t.Error("expected repeat Notes to reuse the slide's notes slide, not create a second")
	}
}
