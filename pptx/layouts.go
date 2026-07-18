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

import "github.com/mmonterroca/pptxgo/drawingml"

// LayoutType names a slide layout's role (p:sldLayout's type attribute,
// ST_SlideLayoutType). Not exhaustive of the full ST_SlideLayoutType set —
// these are the standard layouts New() registers; see newStandardLayouts.
type LayoutType string

// Standard layout types, in the order newStandardLayouts registers them.
const (
	LayoutBlank           LayoutType = "blank"   // no placeholders — pptxgo's original single layout
	LayoutTitleSlide      LayoutType = "title"   // centered title + subtitle
	LayoutTitleAndContent LayoutType = "obj"     // title + one body placeholder
	LayoutSectionHeader   LayoutType = "secHead" // title + body, section-break role
	LayoutTwoContent      LayoutType = "twoObj"  // title + two side-by-side body placeholders
)

// layoutSlide pairs a standard layout's already-built part with the
// LayoutType New() needs to look one up by role (e.g. for AddSlide's
// default in a later phase).
type layoutSlide struct {
	layoutType LayoutType
	xml        *XMLSlideLayout
}

// newStandardLayouts returns pptxgo's fixed catalog of slide layouts, in
// the order they are registered as slideLayoutN.xml parts — index 0 is
// slideLayout1.xml, and so on. LayoutBlank is registered first so it keeps
// occupying slideLayout1.xml, the part AddSlide has always pointed every
// slide at; this catalog only adds new layouts, it does not change what an
// existing AddSlide call renders.
//
// Placeholder geometry is proportioned against the given slide canvas,
// mirroring newMasterSpTree's own proportions. A layout placeholder omits
// a:xfrm (inheriting position from the master) only when its type+idx
// matches one of the master's own two placeholders (title idx=0, body
// idx=1) — see newMasterSpTree. Every other placeholder here (ctrTitle,
// subTitle, and a layout's second body) has no master counterpart to
// inherit from, so it declares its own explicit geometry instead of
// leaving position unresolved.
// newTitleAndSingleBodySpTree builds a title(idx=0)+body(idx=1) placeholder
// spTree that inherits both placeholders' geometry from the master (both
// type+idx pairs match one of the master's own placeholders, so neither
// needs its own a:xfrm) — shared by Title and Content and Section Header,
// which differ only in the body placeholder's display name.
func newTitleAndSingleBodySpTree(bodyName string) *SpTree {
	spTree := NewEmptySpTree()
	spTree.Content = append(spTree.Content,
		newPlaceholderShape(2, "Title Placeholder 2", PlaceholderTitle, 0, nil),
		newPlaceholderShape(3, bodyName, PlaceholderBody, 1, nil),
	)
	return spTree
}

func newStandardLayouts(slideWidthEMU, slideHeightEMU int) []layoutSlide {
	margin := slideWidthEMU * masterMarginPct / 100
	fullW := slideWidthEMU - 2*margin

	blank := NewEmptySpTree()

	titleAndContent := newTitleAndSingleBodySpTree("Content Placeholder 3")
	sectionHeader := newTitleAndSingleBodySpTree("Text Placeholder 3")

	// Title Slide: ctrTitle/subTitle share no type+idx with any master
	// placeholder, so both get their own geometry — centered around the
	// upper third of the canvas.
	ctrTitleH := slideHeightEMU * 20 / 100
	ctrTitleY := slideHeightEMU*35/100 - ctrTitleH/2
	subTitleH := slideHeightEMU * 10 / 100
	subTitleY := ctrTitleY + ctrTitleH + slideHeightEMU*masterTitleGapPct/100
	titleSlide := NewEmptySpTree()
	titleSlide.Content = append(titleSlide.Content,
		newPlaceholderShape(2, "Title Placeholder 2", PlaceholderCtrTitle, 0, &drawingml.Xfrm{
			Off: &drawingml.Off{X: margin, Y: ctrTitleY},
			Ext: &drawingml.Ext{Cx: fullW, Cy: ctrTitleH},
		}),
		newPlaceholderShape(3, "Subtitle Placeholder 3", PlaceholderSubTitle, 1, &drawingml.Xfrm{
			Off: &drawingml.Off{X: margin, Y: subTitleY},
			Ext: &drawingml.Ext{Cx: fullW, Cy: subTitleH},
		}),
	)

	// Two Content: idx=2's second body has no master counterpart either,
	// so both bodies (not just the second) get their own explicit
	// half-width geometry — reusing the master's full-width body
	// placeholder for idx=1 alone would leave it overlapping idx=2 rather
	// than sitting beside it.
	colGap := slideWidthEMU * masterTitleGapPct / 100
	colW := (fullW - colGap) / 2
	_, bodyY, _, bodyH := masterBodyRect(slideWidthEMU, slideHeightEMU)
	twoContent := NewEmptySpTree()
	twoContent.Content = append(twoContent.Content,
		newPlaceholderShape(2, "Title Placeholder 2", PlaceholderTitle, 0, nil),
		newPlaceholderShape(3, "Left Content Placeholder 3", PlaceholderBody, 1, &drawingml.Xfrm{
			Off: &drawingml.Off{X: margin, Y: bodyY},
			Ext: &drawingml.Ext{Cx: colW, Cy: bodyH},
		}),
		newPlaceholderShape(4, "Right Content Placeholder 4", PlaceholderBody, 2, &drawingml.Xfrm{
			Off: &drawingml.Off{X: margin + colW + colGap, Y: bodyY},
			Ext: &drawingml.Ext{Cx: colW, Cy: bodyH},
		}),
	)

	build := func(t LayoutType, spTree *SpTree) layoutSlide {
		return layoutSlide{layoutType: t, xml: &XMLSlideLayout{
			XmlnsA:    drawingml.NamespaceMain,
			XmlnsR:    drawingml.NamespaceRelationships,
			XmlnsP:    NamespaceMain,
			Type:      string(t),
			CSld:      &CSld{SpTree: spTree},
			ClrMapOvr: NewClrMapOvrInherit(),
		}}
	}

	return []layoutSlide{
		build(LayoutBlank, blank),
		build(LayoutTitleSlide, titleSlide),
		build(LayoutTitleAndContent, titleAndContent),
		build(LayoutSectionHeader, sectionHeader),
		build(LayoutTwoContent, twoContent),
	}
}
