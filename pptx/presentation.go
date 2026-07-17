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
	"io"

	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/opc"
)

// Slide 16:9 widescreen canvas size (13.333in x 7.5in) and notes page size
// (7.5in x 10in), in EMUs — PowerPoint's own modern defaults.
const (
	slideWidthEMU  = 12192000
	slideHeightEMU = 6858000
	notesWidthEMU  = 6858000
	notesHeightEMU = 9144000
)

// Conventional ID ranges: PowerPoint reserves the low range of the 32-bit
// ID space and starts slide master IDs at 2^31. Slide IDs and layout IDs
// need only be unique within the presentation; these starting points match
// what every real Office-produced file uses.
const (
	firstSldMasterID = 2147483648
	firstSldLayoutID = 2147483649
	firstSldID       = 256
)

// Presentation is a minimal, single-slide PresentationML document: one
// theme, one slide master, one slide layout, and one blank slide — the
// walking skeleton. It exists to prove the opc.Package plumbing produces a
// file PowerPoint accepts before any real content (text, images, tables)
// is layered on top in later phases.
type Presentation struct {
	pkg *opc.Package
}

// New builds a minimal, single-blank-slide presentation.
func New() *Presentation {
	pkg := opc.NewPackage()

	pkg.AddRawPart(PathTheme1, opc.ContentTypeTheme, []byte(defaultTheme))

	// Add the master's relationships first and thread the generated rIds
	// into the XML, rather than restating "rId2" as a literal that silently
	// depends on the exact order of these Add calls. The master body only
	// references the layout by rId; the theme rel lives in the .rels but is
	// not referenced from the body.
	masterRels := pkg.Relationships(PathSlideMaster1)
	if _, err := masterRels.Add(opc.RelTypeTheme, "../theme/theme1.xml", "Internal"); err != nil {
		panic(err) // static, well-formed arguments; cannot fail
	}
	layoutRID, err := masterRels.Add(RelTypeSlideLayout, "../slideLayouts/slideLayout1.xml", "Internal")
	if err != nil {
		panic(err)
	}
	pkg.AddPart(PathSlideMaster1, ContentTypeSlideMaster, &XMLSlideMaster{
		XmlnsA: drawingml.NamespaceMain,
		XmlnsR: drawingml.NamespaceRelationships,
		XmlnsP: NamespaceMain,
		CSld:   &CSld{SpTree: NewEmptySpTree()},
		ClrMap: NewDefaultClrMap(),
		SldLayoutIdLst: &SldLayoutIdLst{
			Entries: []*SldLayoutId{{ID: firstSldLayoutID, RID: layoutRID}},
		},
		TxStyles: NewDefaultTxStyles(),
	})

	pkg.AddPart(PathSlideLayout1, ContentTypeSlideLayout, &XMLSlideLayout{
		XmlnsA:    drawingml.NamespaceMain,
		XmlnsR:    drawingml.NamespaceRelationships,
		XmlnsP:    NamespaceMain,
		Type:      "blank",
		CSld:      &CSld{SpTree: NewEmptySpTree()},
		ClrMapOvr: NewClrMapOvrInherit(),
	})
	if _, err := pkg.Relationships(PathSlideLayout1).Add(RelTypeSlideMaster, "../slideMasters/slideMaster1.xml", "Internal"); err != nil {
		panic(err)
	}

	slide1 := SlidePath(1)
	pkg.AddPart(slide1, ContentTypeSlide, &XMLSlide{
		XmlnsA:    drawingml.NamespaceMain,
		XmlnsR:    drawingml.NamespaceRelationships,
		XmlnsP:    NamespaceMain,
		CSld:      &CSld{SpTree: NewEmptySpTree()},
		ClrMapOvr: NewClrMapOvrInherit(),
	})
	if _, err := pkg.Relationships(slide1).Add(RelTypeSlideLayout, "../slideLayouts/slideLayout1.xml", "Internal"); err != nil {
		panic(err)
	}

	// Same pattern: add the presentation's relationships and reference the
	// master and slide by the rIds Add returns, not by hardcoded literals.
	presRels := pkg.Relationships(PathPresentation)
	masterRID, err := presRels.Add(RelTypeSlideMaster, "slideMasters/slideMaster1.xml", "Internal")
	if err != nil {
		panic(err)
	}
	slideRID, err := presRels.Add(RelTypeSlide, "slides/slide1.xml", "Internal")
	if err != nil {
		panic(err)
	}
	pkg.AddPart(PathPresentation, ContentTypePresentation, &XMLPresentation{
		XmlnsA: drawingml.NamespaceMain,
		XmlnsR: drawingml.NamespaceRelationships,
		XmlnsP: NamespaceMain,
		SldMasterIdLst: &SldMasterIdLst{
			Entries: []*SldMasterId{{ID: firstSldMasterID, RID: masterRID}},
		},
		SldIdLst: &SldIdLst{
			Entries: []*SldId{{ID: firstSldID, RID: slideRID}},
		},
		SldSz:   &SldSz{Cx: slideWidthEMU, Cy: slideHeightEMU, Type: "screen16x9"},
		NotesSz: &NotesSz{Cx: notesWidthEMU, Cy: notesHeightEMU},
	})

	pkg.AddPart(PathCoreProps, opc.ContentTypeCoreProperties, NewCoreProperties("", ""))
	pkg.AddPart(PathAppProps, opc.ContentTypeExtendedProperties, NewAppProperties())

	rootRels := pkg.Relationships("")
	if _, err := rootRels.Add(RelTypeOfficeDocument, PathPresentation, "Internal"); err != nil {
		panic(err)
	}
	if _, err := rootRels.Add(opc.RelTypeCoreProperties, PathCoreProps, "Internal"); err != nil {
		panic(err)
	}
	if _, err := rootRels.Add(opc.RelTypeExtendedProperties, PathAppProps, "Internal"); err != nil {
		panic(err)
	}

	return &Presentation{pkg: pkg}
}

// Save writes the presentation to w as a .pptx file.
func (p *Presentation) Save(w io.Writer) error {
	return p.pkg.Write(w)
}
