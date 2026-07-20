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

	"github.com/mmonterroca/pptxgo/drawingml"
)

// XMLPresentation represents ppt/presentation.xml (p:presentation).
type XMLPresentation struct {
	XMLName xml.Name `xml:"p:presentation"`
	XmlnsA  string   `xml:"xmlns:a,attr"`
	XmlnsR  string   `xml:"xmlns:r,attr"`
	XmlnsP  string   `xml:"xmlns:p,attr"`
	// Schema order (CT_Presentation): sldMasterIdLst, notesMasterIdLst,
	// handoutMasterIdLst, sldIdLst, sldSz, notesSz. NotesMasterIdLst is nil
	// until the first speaker note is added (see Slide.Notes), so a deck with
	// no notes emits none.
	SldMasterIdLst   *SldMasterIdLst   `xml:"p:sldMasterIdLst"`
	NotesMasterIdLst *NotesMasterIdLst `xml:"p:notesMasterIdLst,omitempty"`
	SldIdLst         *SldIdLst         `xml:"p:sldIdLst"`
	SldSz            *SldSz            `xml:"p:sldSz"`
	NotesSz          *NotesSz          `xml:"p:notesSz"`
}

// NotesMasterIdLst is p:notesMasterIdLst, the (at most one, for pptxgo) list
// of notes masters.
type NotesMasterIdLst struct {
	XMLName xml.Name         `xml:"p:notesMasterIdLst"`
	Entries []*NotesMasterId `xml:"p:notesMasterId"`
}

// NotesMasterId is a single p:notesMasterId entry, referencing the notes
// master part via relationship ID.
type NotesMasterId struct {
	XMLName xml.Name `xml:"p:notesMasterId"`
	RID     string   `xml:"r:id,attr"`
}

// SldMasterIdLst is p:sldMasterIdLst, the list of slide masters.
type SldMasterIdLst struct {
	XMLName xml.Name       `xml:"p:sldMasterIdLst"`
	Entries []*SldMasterId `xml:"p:sldMasterId"`
}

// SldMasterId is a single p:sldMasterId entry, referencing a slideMaster
// part via relationship ID.
type SldMasterId struct {
	XMLName xml.Name `xml:"p:sldMasterId"`
	ID      uint32   `xml:"id,attr"`
	RID     string   `xml:"r:id,attr"`
}

// SldIdLst is p:sldIdLst, the ordered list of slides in the presentation.
type SldIdLst struct {
	XMLName xml.Name `xml:"p:sldIdLst"`
	Entries []*SldId `xml:"p:sldId"`
}

// SldId is a single p:sldId entry, referencing a slide part via relationship ID.
type SldId struct {
	XMLName xml.Name `xml:"p:sldId"`
	ID      uint32   `xml:"id,attr"`
	RID     string   `xml:"r:id,attr"`
}

// SldSz is p:sldSz, the slide canvas size in EMUs.
type SldSz struct {
	XMLName xml.Name `xml:"p:sldSz"`
	Cx      int      `xml:"cx,attr"`
	Cy      int      `xml:"cy,attr"`
	Type    string   `xml:"type,attr,omitempty"`
}

// NotesSz is p:notesSz, the notes page canvas size in EMUs.
type NotesSz struct {
	XMLName xml.Name `xml:"p:notesSz"`
	Cx      int      `xml:"cx,attr"`
	Cy      int      `xml:"cy,attr"`
}

// XMLSlide represents a ppt/slides/slideN.xml part (p:sld).
type XMLSlide struct {
	XMLName   xml.Name   `xml:"p:sld"`
	XmlnsA    string     `xml:"xmlns:a,attr"`
	XmlnsR    string     `xml:"xmlns:r,attr"`
	XmlnsP    string     `xml:"xmlns:p,attr"`
	CSld      *CSld      `xml:"p:cSld"`
	ClrMapOvr *ClrMapOvr `xml:"p:clrMapOvr"`
}

// CSld is p:cSld, the common slide data container: an optional background
// followed by the shape tree (plus, on slides only, an optional name
// attribute). Bg comes before SpTree in the struct because CT_CommonSlideData
// requires it there — bg is minOccurs=0, but when present it must precede
// the (always-required) spTree.
type CSld struct {
	XMLName xml.Name `xml:"p:cSld"`
	Bg      *Bg      `xml:"p:bg,omitempty"`
	SpTree  *SpTree  `xml:"p:spTree"`
}

// Bg is p:bg (CT_Background): a slide's own background, overriding
// whatever its layout/master would otherwise supply. Only the simplest
// path — an explicit fill via BgPr — is modeled; bgRef (a reference into
// the theme's format-scheme background styles) is out of scope.
type Bg struct {
	XMLName xml.Name `xml:"p:bg"`
	BgPr    *BgPr    `xml:"p:bgPr"`
}

// BgPr is p:bgPr (CT_BackgroundProperties): the background's own fill.
// Fill and Gradient are the schema's EG_FillProperties choice: at most one
// should ever be set — Slide.Background/BackgroundScheme/BackgroundGradient
// enforce that by clearing the other whenever one is set.
type BgPr struct {
	XMLName  xml.Name             `xml:"p:bgPr"`
	Fill     *drawingml.SolidFill `xml:"a:solidFill,omitempty"`
	Gradient *drawingml.GradFill  `xml:"a:gradFill,omitempty"`
}

// SpTree is p:spTree, the shape tree: the root container for every visible
// shape on a slide, layout, or master. Content is left as `any` — text
// boxes, pictures, and graphic frames are modeled by later phases; the
// walking skeleton only needs an empty, schema-valid tree.
type SpTree struct {
	XMLName   xml.Name   `xml:"p:spTree"`
	NvGrpSpPr *NvGrpSpPr `xml:"p:nvGrpSpPr"`
	GrpSpPr   *GrpSpPr   `xml:"p:grpSpPr"`
	Content   []any      `xml:",any"`
}

// NvGrpSpPr is p:nvGrpSpPr, the shape tree's own (required, always-empty
// group) non-visual properties.
type NvGrpSpPr struct {
	XMLName    xml.Name `xml:"p:nvGrpSpPr"`
	CNvPr      *CNvPr   `xml:"p:cNvPr"`
	CNvGrpSpPr *struct {
		XMLName xml.Name `xml:"p:cNvGrpSpPr"`
	} `xml:"p:cNvGrpSpPr"`
	NvPr *struct {
		XMLName xml.Name `xml:"p:nvPr"`
	} `xml:"p:nvPr"`
}

// CNvPr is p:cNvPr, the non-visual drawing properties (ID and name) shared
// by every shape-like element.
type CNvPr struct {
	XMLName xml.Name `xml:"p:cNvPr"`
	ID      uint32   `xml:"id,attr"`
	Name    string   `xml:"name,attr"`
}

// GrpSpPr is p:grpSpPr (CT_GroupShapeProperties): a group shape's own
// properties. Xfrm is nil for the slide's root spTree (the walking
// skeleton's own usage — the root group never needs a non-identity child
// space) and set for a nested p:grpSp (see Slide.AddGroup), the only two
// contexts this type is used in. Fill/effect/3d properties are out of
// scope — a group's own visual properties are rarely set directly; its
// member shapes carry their own.
type GrpSpPr struct {
	XMLName xml.Name             `xml:"p:grpSpPr"`
	Xfrm    *drawingml.GroupXfrm `xml:"a:xfrm,omitempty"`
}

// CNvSpPr is p:cNvSpPr, non-visual drawing properties specific to shapes.
// TxBox marks the shape as a text box rather than an auto-shape — required
// for PowerPoint to treat a bare rectangle as a text container.
type CNvSpPr struct {
	XMLName xml.Name `xml:"p:cNvSpPr"`
	TxBox   bool     `xml:"txBox,attr,omitempty"`
}

// NvPr is p:nvPr: placeholder-linkage information shared by every shape's
// non-visual properties. Ph is nil for an ordinary shape; set, it marks
// this shape as a placeholder — see Ph.
type NvPr struct {
	XMLName xml.Name `xml:"p:nvPr"`
	Ph      *Ph      `xml:"p:ph,omitempty"`
}

// Ph is p:ph (CT_Placeholder): marks a shape as a placeholder, linking it
// by Type+Idx to the correspondingly-typed placeholder in this part's
// layout (and, from there, the master) for position/formatting
// inheritance — a slide (or layout) placeholder that sets no a:xfrm of its
// own inherits the layout's (or master's). Idx is the schema's own
// ST_PlaceholderIndex default of 0 when unset (plain uint32 + omitempty,
// not *int — unlike MarL/Lvl/Indent elsewhere, 0 here is genuinely "not
// set, use the default" rather than a meaningful explicit value), so a
// title or single-body placeholder never needs to set it; only a second
// placeholder of the same type on one slide (e.g. a two-content layout's
// second body) needs a distinct idx. uint32, not int: ST_PlaceholderIndex
// is xsd:unsignedInt, so a negative value would be schema-invalid — the
// type itself rules that out rather than needing a runtime check.
type Ph struct {
	XMLName xml.Name        `xml:"p:ph"`
	Type    PlaceholderType `xml:"type,attr,omitempty"`
	Idx     uint32          `xml:"idx,attr,omitempty"`
}

// ClrMapOvr is p:clrMapOvr, present on every slide and slide layout: it
// either inherits the owning master's color map verbatim or overrides it.
// The walking skeleton always inherits.
type ClrMapOvr struct {
	XMLName          xml.Name `xml:"p:clrMapOvr"`
	MasterClrMapping *struct {
		XMLName xml.Name `xml:"a:masterClrMapping"`
	} `xml:"a:masterClrMapping"`
}

// NewClrMapOvrInherit returns a ClrMapOvr that inherits the master's color map.
func NewClrMapOvrInherit() *ClrMapOvr {
	return &ClrMapOvr{MasterClrMapping: &struct {
		XMLName xml.Name `xml:"a:masterClrMapping"`
	}{}}
}

// newNvGrpSpPr builds a p:nvGrpSpPr (CT_GroupShapeNonVisual) with the given
// id and name — shared by NewEmptySpTree (the slide's own root group,
// id=1, unnamed) and Slide.AddGroup (a nested p:grpSp, an allocated id and
// a "Group N" name). Both contexts are otherwise identical: an always-empty
// cNvGrpSpPr and nvPr.
func newNvGrpSpPr(id uint32, name string) *NvGrpSpPr {
	return &NvGrpSpPr{
		CNvPr: &CNvPr{ID: id, Name: name},
		CNvGrpSpPr: &struct {
			XMLName xml.Name `xml:"p:cNvGrpSpPr"`
		}{},
		NvPr: &struct {
			XMLName xml.Name `xml:"p:nvPr"`
		}{},
	}
}

// NewEmptySpTree returns a minimal, schema-valid shape tree with no shapes.
func NewEmptySpTree() *SpTree {
	return &SpTree{
		NvGrpSpPr: newNvGrpSpPr(1, ""),
		GrpSpPr:   &GrpSpPr{},
	}
}
