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

import "encoding/xml"

// XMLPresentation represents ppt/presentation.xml (p:presentation).
type XMLPresentation struct {
	XMLName        xml.Name        `xml:"p:presentation"`
	XmlnsA         string          `xml:"xmlns:a,attr"`
	XmlnsR         string          `xml:"xmlns:r,attr"`
	XmlnsP         string          `xml:"xmlns:p,attr"`
	SldMasterIdLst *SldMasterIdLst `xml:"p:sldMasterIdLst"`
	SldIdLst       *SldIdLst       `xml:"p:sldIdLst"`
	SldSz          *SldSz          `xml:"p:sldSz"`
	NotesSz        *NotesSz        `xml:"p:notesSz"`
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

// CSld is p:cSld, the common slide data container (shape tree plus, on
// slides only, an optional name attribute).
type CSld struct {
	XMLName xml.Name `xml:"p:cSld"`
	SpTree  *SpTree  `xml:"p:spTree"`
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

// GrpSpPr is p:grpSpPr, the shape tree's group shape properties. Empty for
// the walking skeleton.
type GrpSpPr struct {
	XMLName xml.Name `xml:"p:grpSpPr"`
}

// CNvSpPr is p:cNvSpPr, non-visual drawing properties specific to shapes.
// TxBox marks the shape as a text box rather than an auto-shape — required
// for PowerPoint to treat a bare rectangle as a text container.
type CNvSpPr struct {
	XMLName xml.Name `xml:"p:cNvSpPr"`
	TxBox   bool     `xml:"txBox,attr,omitempty"`
}

// NvPr is p:nvPr: placeholder-linkage information shared by every shape's
// non-visual properties. Always empty until placeholders (p:ph) land in a
// later phase.
type NvPr struct {
	XMLName xml.Name `xml:"p:nvPr"`
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

// NewEmptySpTree returns a minimal, schema-valid shape tree with no shapes.
func NewEmptySpTree() *SpTree {
	return &SpTree{
		NvGrpSpPr: &NvGrpSpPr{
			CNvPr: &CNvPr{ID: 1, Name: ""},
			CNvGrpSpPr: &struct {
				XMLName xml.Name `xml:"p:cNvGrpSpPr"`
			}{},
			NvPr: &struct {
				XMLName xml.Name `xml:"p:nvPr"`
			}{},
		},
		GrpSpPr: &GrpSpPr{},
	}
}
