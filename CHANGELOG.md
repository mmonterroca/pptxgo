# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Added

- `Presentation.AddSlide` to append slides to a presentation; `New()` now
  starts with zero slides instead of hardcoding one.
- `Slide.AddTextBox` and a fluent `TextBox`/`Paragraph` builder
  (`AddParagraph`, `Text`, `Bold`, `Italic`, `Underline`, `FontSize`,
  `Font`, `Color`, `Alignment`) for adding formatted text to a slide.
- `drawingml` text model (`TextBody`, `Paragraph`, `Run`, `RPr`, `Latin`)
  for `a:txBody`/`a:p`/`a:r`/`a:rPr`/`a:t`.
- `pptx.Inches`, `pptx.Points`, `pptx.Emu`, `pptx.RGB`, and alignment
  constants (`AlignLeft`/`AlignCenter`/`AlignRight`/`AlignJustify`).
- `Slide.AddImage`, `AddImageWithSize`, `AddImageFromBytes`, and
  `AddImageFromBytesWithSize` to embed PNG/JPEG/GIF images as `p:pic`, with
  automatic format and pixel-dimension detection (via `image.DecodeConfig`,
  no new dependency) and 96 DPI auto-sizing when no explicit size is given.
- `TextBox.Fill`/`TextBox.Border` and `PictureRef.Border` for a solid shape
  background and outline (`a:solidFill`/`a:ln`) on text boxes and images.
- `drawingml.Ln` (`a:ln`) for solid-color shape/picture outlines.

### Changed

- `Presentation.Save` now returns the first error accumulated by a builder
  call (e.g. an out-of-range `FontSize`) instead of every method needing its
  own error return.

## [0.1.0]

- Walking skeleton: `opc` packaging layer, `drawingml` shared primitives,
  and `pptx.New()` producing a minimal, single-blank-slide presentation.
