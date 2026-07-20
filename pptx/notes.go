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
	"strconv"

	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/opc"
)

// Notes-page body placeholder geometry, in EMUs: the lower portion of the
// 7.5in x 10in notes page (notesWidthEMU x notesHeightEMU), leaving room a
// real notes page uses for the slide thumbnail above it. Speaker notes are
// read from this body placeholder regardless of its exact rect, so these are
// conventional rather than load-bearing.
const (
	notesBodyX = 457200  // 0.5in left margin
	notesBodyY = 4114800 // ~4.5in from the top (below the slide-image area)
	notesBodyW = 5943600 // page width minus 0.5in margins
	notesBodyH = 4572000 // 5in tall
)

// XMLNotesMaster represents ppt/notesMasters/notesMaster1.xml (p:notesMaster):
// the single notes master every notes slide inherits from. pptxgo models a
// minimal one — a notes body placeholder and the standard color map — created
// lazily the first time any slide gets speaker notes.
type XMLNotesMaster struct {
	XMLName xml.Name `xml:"p:notesMaster"`
	XmlnsA  string   `xml:"xmlns:a,attr"`
	XmlnsR  string   `xml:"xmlns:r,attr"`
	XmlnsP  string   `xml:"xmlns:p,attr"`
	CSld    *CSld    `xml:"p:cSld"`
	ClrMap  *ClrMap  `xml:"p:clrMap"`
}

// XMLNotesSlide represents a ppt/notesSlides/notesSlideN.xml part (p:notes):
// one slide's speaker notes, held in a body placeholder.
type XMLNotesSlide struct {
	XMLName   xml.Name   `xml:"p:notes"`
	XmlnsA    string     `xml:"xmlns:a,attr"`
	XmlnsR    string     `xml:"xmlns:r,attr"`
	XmlnsP    string     `xml:"xmlns:p,attr"`
	CSld      *CSld      `xml:"p:cSld"`
	ClrMapOvr *ClrMapOvr `xml:"p:clrMapOvr"`
}

// newNotesBodyPlaceholder builds the notes body placeholder (a p:sp with a
// type="body" idx="1" p:ph) with its own geometry, so it renders on the notes
// page without depending on inheriting a rect from the notes master.
func newNotesBodyPlaceholder() *Shape {
	body := newPlaceholderShape(2, "Notes Placeholder 2", PlaceholderBody, 1, &drawingml.Xfrm{
		Off: &drawingml.Off{X: notesBodyX, Y: notesBodyY},
		Ext: &drawingml.Ext{Cx: notesBodyW, Cy: notesBodyH},
	})
	body.SpPr.PrstGeom = &drawingml.PrstGeom{Prst: string(ShapeRect), AvLst: &drawingml.AvLst{}}
	return body
}

// ensureNotesMaster creates the presentation's single notes master part (and
// its theme relationship, and the presentation's notesMasterIdLst entry) the
// first time any slide gets notes; later calls are no-ops. A deck with no
// notes never gets a notes master, so its output is unchanged.
func (p *Presentation) ensureNotesMaster() {
	if p.notesMasterCreated {
		return
	}

	spTree := NewEmptySpTree()
	spTree.Content = append(spTree.Content, newNotesBodyPlaceholder())

	p.pkg.AddPart(PathNotesMaster1, ContentTypeNotesMaster, &XMLNotesMaster{
		XmlnsA: drawingml.NamespaceMain,
		XmlnsR: drawingml.NamespaceRelationships,
		XmlnsP: NamespaceMain,
		CSld:   &CSld{SpTree: spTree},
		ClrMap: NewDefaultClrMap(),
	})

	// A master part references a theme; the notes master shares the one theme.
	if _, err := p.pkg.Relationships(PathNotesMaster1).Add(opc.RelTypeTheme, "../theme/theme1.xml", "Internal"); err != nil {
		panic(err) // static, well-formed arguments; cannot fail
	}

	rid, err := p.presRels.Add(RelTypeNotesMaster, "notesMasters/notesMaster1.xml", "Internal")
	if err != nil {
		panic(err)
	}
	p.pres.NotesMasterIdLst = &NotesMasterIdLst{Entries: []*NotesMasterId{{RID: rid}}}
	p.notesMasterCreated = true
}

// Notes sets the slide's speaker notes to text — the note that appears in
// PowerPoint's notes pane and on the printed notes page. Embedded newlines
// ("\n") become line breaks within the note. Calling Notes again on the same
// slide appends the new text as a further paragraph rather than replacing it.
//
// The first call across the whole presentation lazily creates the single
// notes master; a deck that never calls Notes emits no notes parts at all.
func (s *Slide) Notes(text string) *Slide {
	// Repeat call: append another paragraph to this slide's existing notes.
	if s.notesBody != nil {
		para := &drawingml.Paragraph{}
		s.notesBody.Paragraphs = append(s.notesBody.Paragraphs, para)
		(&Paragraph{pres: s.pres, slidePath: NotesSlidePath(s.num), p: para}).Text(text)
		return s
	}

	s.pres.ensureNotesMaster()

	notesPath := NotesSlidePath(s.num)
	body := newNotesBodyPlaceholder()

	para := &drawingml.Paragraph{}
	body.TxBody.Paragraphs = append(body.TxBody.Paragraphs, para)
	(&Paragraph{pres: s.pres, slidePath: notesPath, p: para}).Text(text)

	spTree := NewEmptySpTree()
	spTree.Content = append(spTree.Content, body)

	s.pres.pkg.AddPart(notesPath, ContentTypeNotesSlide, &XMLNotesSlide{
		XmlnsA:    drawingml.NamespaceMain,
		XmlnsR:    drawingml.NamespaceRelationships,
		XmlnsP:    NamespaceMain,
		CSld:      &CSld{SpTree: spTree},
		ClrMapOvr: NewClrMapOvrInherit(),
	})

	// The notes slide references its notes master and the slide it annotates.
	nrels := s.pres.pkg.Relationships(notesPath)
	if _, err := nrels.Add(RelTypeNotesMaster, "../notesMasters/notesMaster1.xml", "Internal"); err != nil {
		panic(err) // static, well-formed arguments; cannot fail
	}
	if _, err := nrels.Add(RelTypeSlide, "../slides/slide"+strconv.Itoa(s.num)+".xml", "Internal"); err != nil {
		panic(err)
	}

	// The slide references its notes slide.
	if _, err := s.pres.pkg.Relationships(s.path).Add(RelTypeNotesSlide, "../notesSlides/notesSlide"+strconv.Itoa(s.num)+".xml", "Internal"); err != nil {
		panic(err)
	}

	s.notesBody = body.TxBody
	return s
}
