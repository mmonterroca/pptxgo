# pptxgo

Microsoft PowerPoint .pptx (OOXML / PresentationML) generation in Go.

[![Go Reference](https://pkg.go.dev/badge/github.com/mmonterroca/pptxgo.svg)](https://pkg.go.dev/github.com/mmonterroca/pptxgo)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

## Status

**Early development.** The OPC packaging layer and a minimal walking-skeleton
presentation (empty slide, valid theme, master/layout) are being built first;
text, images, tables, and templates follow.

## Design

The core insight driving this project's architecture: PPTX content is
DrawingML. Every slide is a `p:spTree` of shapes, and all text lives in
`a:txBody / a:p / a:r / a:t`. That put a hard constraint on the packaging
layer from day one — it had to be **part-centric**: a package is a map of
parts (path → content-type → bytes-or-struct), never a hardcoded sequence of
named files. That's what makes "open an existing .pptx template and only
replace the parts you touch" the same code path as "generate everything from
scratch" — not a bolted-on round-trip mode added later.

Two packages are written to be extraction-ready from the start:

- `opc/` — the Open Packaging Conventions layer (parts, relationships,
  content-types, ZIP serialization). Format-agnostic; knows nothing about
  slides or paragraphs.
- `drawingml/` — the DrawingML primitives shared by DOCX/PPTX/XLSX (`a:xfrm`,
  `a:off`, `a:ext`, `a:blip`, `a:prstGeom`, colors, transforms). It only ever
  emits the `a:` namespace — the picture container that wraps a blip fill
  differs by host format (`pic:pic` in a Word-embedded graphic, `p:pic` on a
  PPTX slide), so that wrapper is deliberately left to the package that
  needs it, built out of these shared primitives.

If a future sibling project needs the same OPC engine, both are designed to
be lifted into a standalone module without a rewrite.

## License

MIT
