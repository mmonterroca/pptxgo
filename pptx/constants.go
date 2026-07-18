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

// Package pptx assembles PresentationML content on top of opc.Package and
// drawingml's shared primitives: presentation.xml, slide masters, slide
// layouts, slides, and the theme.
package pptx

import (
	"strconv"

	"github.com/mmonterroca/pptxgo/opc"
)

// NamespaceMain is the PresentationML main namespace ("p:").
const NamespaceMain = "http://schemas.openxmlformats.org/presentationml/2006/main"

// Part paths, relative to the package root.
const (
	PathPresentation = "ppt/presentation.xml"
	PathTheme1       = "ppt/theme/theme1.xml"
	PathSlideMaster1 = "ppt/slideMasters/slideMaster1.xml"
	PathCoreProps    = "docProps/core.xml"
	PathAppProps     = "docProps/app.xml"

	// PathSlideLayout1 is SlideLayoutPath(1) — the always-present LayoutBlank
	// part — spelled as a constant for callers that want the blank layout's
	// path without computing it.
	PathSlideLayout1 = "ppt/slideLayouts/slideLayout1.xml"
)

// SlidePath returns the part path for the nth slide (1-indexed).
func SlidePath(n int) string {
	return "ppt/slides/slide" + strconv.Itoa(n) + ".xml"
}

// SlideLayoutPath returns the part path for the nth slide layout
// (1-indexed) — see newStandardLayouts for what each index holds.
func SlideLayoutPath(n int) string {
	return "ppt/slideLayouts/slideLayout" + strconv.Itoa(n) + ".xml"
}

// slideLayoutRelTarget returns the nth slide layout's path as a
// relationship target relative to another ppt/ subdirectory's own part
// (e.g. ppt/slideMasters/slideMaster1.xml or ppt/slides/slideN.xml) — the
// form both New (the master's own layout rels) and AddSlide (a slide's
// layout rel) need, as opposed to SlideLayoutPath's package-root-relative
// form used for registering the part itself.
func slideLayoutRelTarget(n int) string {
	return "../slideLayouts/slideLayout" + strconv.Itoa(n) + ".xml"
}

// Content types specific to PresentationML.
const (
	ContentTypePresentation = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
	ContentTypeSlideMaster  = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
	ContentTypeSlideLayout  = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
	ContentTypeSlide        = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
)

// Relationship types specific to PresentationML (image/theme/hyperlink and
// the OPC-level metadata rel types are shared and live in opc.RelType*).
const (
	RelTypeOfficeDocument = opc.NamespaceOfficeDocumentRels + "/officeDocument"
	RelTypeSlideMaster    = opc.NamespaceOfficeDocumentRels + "/slideMaster"
	RelTypeSlideLayout    = opc.NamespaceOfficeDocumentRels + "/slideLayout"
	RelTypeSlide          = opc.NamespaceOfficeDocumentRels + "/slide"
)
