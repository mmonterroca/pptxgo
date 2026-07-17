# pptxgo

Microsoft PowerPoint .pptx (OOXML / PresentationML) generation in Go.

[![Go Reference](https://pkg.go.dev/badge/github.com/mmonterroca/pptxgo.svg)](https://pkg.go.dev/github.com/mmonterroca/pptxgo)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

## Status

**Early development.** The OPC packaging layer and drawingml primitives are
in place, and `pptx.New()` produces a minimal, single-blank-slide
presentation — verified against both the Open XML SDK's schema validator
and LibreOffice Impress (see Verification below). Text, images, tables, and
templates follow.

```go
p := pptx.New()
f, _ := os.Create("presentation.pptx")
p.Save(f)
```

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

## Verification

A file that unzips fine can still be schema-invalid, and Go-level tests
alone can't tell you that — they only check what you thought to assert.
`make check` runs three layers:

1. `go test ./...` — structural regression tests (every relationship target
   resolves to a part that exists, `[Content_Types].xml` covers every part,
   no duplicate relationship IDs within an owner).
2. `PptxValidator/` — an `OpenXmlValidator` (DocumentFormat.OpenXml, the
   same library Microsoft ships) run against a generated demo file. Wired
   into CI from the first commit — unlike docxgo, where the equivalent
   validator existed for nine months before CI ever invoked it.
3. Opening the file in a real consumer. `make validate` also works as a
   smoke test if you pipe its output pptx through
   `soffice --headless --convert-to pdf`; a manual open in PowerPoint
   itself remains the authoritative check no automated tool replaces.

```
make check     # test + build + generate + OpenXML SDK validation
```

## License

MIT
