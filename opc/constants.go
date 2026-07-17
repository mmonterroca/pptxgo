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

// Package opc implements the Open Packaging Conventions (ECMA-376 Part 2):
// the ZIP container, [Content_Types].xml, and relationship (.rels) parts
// shared by every OOXML format (DOCX, PPTX, XLSX). It knows nothing about
// PresentationML, WordprocessingML, or any other content schema — a Package
// is just a map of parts (path, content type, and either raw bytes or a
// value to marshal), plus the relationships between them.
//
// This separation is deliberate: it is what lets a package be built by
// generating every part from scratch, by loading an existing file and
// replacing only the parts that change, or any mix of the two, through the
// exact same code path.
package opc

// OPC namespaces (package-level; format-specific namespaces such as
// PresentationML's "p:" live in the consuming package, not here).
const (
	// NamespacePackageRels is the namespace for relationship (.rels) parts.
	NamespacePackageRels = "http://schemas.openxmlformats.org/package/2006/relationships"

	// NamespaceContentTypes is the namespace for [Content_Types].xml.
	NamespaceContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types"

	// NamespaceOfficeDocumentRels is the namespace prefix for officeDocument
	// relationship types (image, hyperlink, and format-specific document
	// relationships are all suffixed onto this).
	NamespaceOfficeDocumentRels = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

	// NamespaceCoreProperties is the namespace for docProps/core.xml.
	NamespaceCoreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"

	// NamespaceExtendedProperties is the namespace for docProps/app.xml.
	NamespaceExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

	// NamespaceDC is the Dublin Core namespace used in core properties.
	NamespaceDC = "http://purl.org/dc/elements/1.1/"

	// NamespaceDCTerms is the Dublin Core Terms namespace used in core properties.
	NamespaceDCTerms = "http://purl.org/dc/terms/"

	// NamespaceXSI is the XML Schema instance namespace used in core properties.
	NamespaceXSI = "http://www.w3.org/2001/XMLSchema-instance"
)

// OPC-level relationship types (shared across DOCX/PPTX/XLSX). Format-specific
// relationship types (e.g. PresentationML's slide/slideLayout/slideMaster)
// live in the consuming package.
const (
	RelTypeImage              = NamespaceOfficeDocumentRels + "/image"
	RelTypeHyperlink          = NamespaceOfficeDocumentRels + "/hyperlink"
	RelTypeTheme              = NamespaceOfficeDocumentRels + "/theme"
	RelTypeCoreProperties     = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	RelTypeExtendedProperties = NamespaceOfficeDocumentRels + "/extended-properties"
)

// OPC-level content types.
const (
	ContentTypeRelationships      = "application/vnd.openxmlformats-package.relationships+xml"
	ContentTypeXML                = "application/xml"
	ContentTypeCoreProperties     = "application/vnd.openxmlformats-package.core-properties+xml"
	ContentTypeExtendedProperties = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
	ContentTypeTheme              = "application/vnd.openxmlformats-officedocument.theme+xml"

	ContentTypePNG  = "image/png"
	ContentTypeJPEG = "image/jpeg"
	ContentTypeGIF  = "image/gif"
	ContentTypeBMP  = "image/bmp"
	ContentTypeTIFF = "image/tiff"
	ContentTypeWMF  = "image/x-wmf"
	ContentTypeEMF  = "image/x-emf"
)

// Fixed OPC part paths.
const (
	// PathContentTypes is the path of the package-wide content types part.
	PathContentTypes = "[Content_Types].xml"

	// PathRootRels is the path of the package-wide root relationships part.
	PathRootRels = "_rels/.rels"
)

// DefaultRelCapacity is the initial capacity for a part's relationship map,
// to reduce allocations for the common small-fan-out case.
const DefaultRelCapacity = 32
