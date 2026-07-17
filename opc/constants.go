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

// Measurement conversions shared by all OOXML DrawingML-based formats.
const (
	// TwipsPerInch is the number of twips in one inch (1440).
	// A twip is 1/1440 of an inch, or 1/20 of a point. Used by WordprocessingML.
	TwipsPerInch = 1440

	// TwipsPerPoint is the number of twips in one point (20).
	TwipsPerPoint = 20

	// PointsPerInch is 72 points per inch.
	PointsPerInch = 72

	// EMUsPerInch is the number of English Metric Units in one inch (914400).
	// EMUs are the unit DrawingML uses for shape transforms (a:off, a:ext).
	EMUsPerInch = 914400

	// EMUsPerTwip is EMUsPerInch / TwipsPerInch (635).
	EMUsPerTwip = 635

	// EMUsPerPoint is EMUsPerInch / PointsPerInch (12700). PresentationML
	// needs this for line widths (a:ln w=) and any point-based positioning;
	// WordprocessingML's half-point/twip units never required it.
	EMUsPerPoint = 12700

	// EMUsPerCentimeter is EMUsPerInch / 2.54, rounded (360000).
	EMUsPerCentimeter = 360000
)

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

// Default capacities for slices/maps, to reduce allocations.
const (
	DefaultRelCapacity   = 32
	DefaultMediaCapacity = 16
)
