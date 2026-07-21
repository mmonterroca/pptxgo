# pptxgo Roadmap

`pptxgo` is a zero-external-dependency Go library for generating and
editing Microsoft PowerPoint `.pptx` (OOXML/PresentationML) files. This
document tracks the master plan against what has actually shipped, so it
can be read on its own — no chat history required.

**Status legend:** ✅ done and merged to `master` · 🚧 in progress ·
⏳ planned, not started · ⏸️ deliberately deferred (see reason).

Every shipped feature below passed this repo's three-layer verification
discipline before merging: (1) `go test ./...` + the OpenXML SDK schema
validator (`make check`), (2) a multi-agent `/code-review` pass with all
confirmed findings fixed, (3) a visual check in LibreOffice and/or real
Microsoft PowerPoint — the only way some defects (a missing reflection
`sy`, a connector's first-paint routing) were ever caught, since
schema-valid XML can still render wrong.

---

## Phase 1 — Packager & walking skeleton ✅

OPC (Open Packaging Conventions) part-centric packaging layer, content
types, relationships, and a minimal end-to-end `.pptx` that opens in
PowerPoint.

## Phase 2 — Text ✅

`p:sp` + `a:txBody` shapes, fluent `AddTextBox`/`AddParagraph` builder
(bold/italic/underline/font/size/color/alignment).

## Phase 3 — Media ✅

`p:pic` image embedding (PNG/JPEG/GIF, auto format + size detection via
EXIF), plus shape fill/border that this repo folded into the same phase.

## Phase 4 — Tables ✅

`a:tbl`/`a:tr`/`a:tc`/`a:tblGrid` inside `p:graphicFrame`, `Slide.AddTable`,
default Office table style. Cell merging (`MergeCells`), per-side cell
borders, and cell fill/anchor followed later (see Phase 8 / completion
batch below) — the base table primitive shipped here.

## Phase 5 — Masters, layouts, placeholders ✅

Master → layout → slide inheritance: `p:ph` placeholders, 5 built-in
slide layouts (Blank, Title Slide, Title and Content, Section Header, Two
Content), `AddSlide(WithLayout(...))`, `Slide.AddPlaceholder`/`Title`/
`Body`. Scoped deliberately as **writer-only** — pptxgo emits
type/idx-matched placeholders and leaves inheritance *resolution* to
PowerPoint/LibreOffice; a resolver is a reader concern, not modeled here.

## Phase 6 — Reader / templates ✅

`pptx.Open`/`OpenBytes`/`OpenReader` hydrate an existing `.pptx` (untouched
parts preserved byte-for-byte), slide navigation, and `{{placeholder}}`
text substitution/merge with run-consolidation (healing PowerPoint's
autocorrect-driven run splitting). See `examples/02_read_and_modify`.

## Phase 7 — Charts (`c:`) ⏳

**Not started.** No `c:chart` element modeling exists yet — only a
`GraphicDataURIChart` constant marking the `graphicFrame` extension point
charts would eventually use. This is the largest remaining item on the
original 7-phase plan.

---

## Phase 8 — Theming & deck chrome ✅

- **Brand theme** (`pptx.Theme`/`ThemeColors`/`ThemeFonts`, `WithTheme`,
  `themes` subpackage with Office/Corporate/Modern presets) — `clrScheme`
  + `fontScheme` modeled as structs instead of a hardcoded literal.
- **Gradient scheme-color stops, table cell fill/anchor, shape
  `Adjust(name, val)`, numbered-bullet start value.**
- **Document metadata** (`pptx.Metadata`, `WithTitle`/`WithAuthor`/
  `WithSubject`/`WithKeywords`/`WithDescription`/`WithCompany`).
- **Speaker notes** (`Slide.Notes`) — lazy notesMaster + per-slide
  notesSlide parts.
- **Footer / date / slide number** (`Slide.Footer`/`DateText`/
  `SlideNumber`) — self-positioned placeholders that don't touch the
  master or layouts, so no Phase 5 regression risk.

## Completion batch — modeled-but-incomplete surfaces ✅

Rounded out four features that existed structurally but had no builder:

- **Per-side table-cell borders** (`TableCell.Border`/`BorderScheme`, 6
  sides).
- **Gradient stop tint/shade** (`Tint`/`Shade`/`Alpha`/`LumMod`/`LumOff`
  color-transform children on `SrgbClr`/`SchemeClr`).
- **Line caps, joins, arrowheads** (`ShapeRef.LineCap`/`LineJoin`/
  `ArrowStart`/`ArrowEnd`).
- **Shape effects** (`a:effectLst`: `Shadow`/`Glow`/`Reflection`/
  `SoftEdges`, shared by shapes and pictures).

Caught in real PowerPoint (not by the SDK validator or code review): the
reflection effect rendered nothing without an explicit `sy="-100000"`
mirror-flip attribute — fixed and now covered by a regression test.

## Native diagrams — grouped shapes + bound connectors ✅

- **`Slide.AddGroup`/`Group.AddShape`/`Group.AddTextBox`** (`p:grpSp`) —
  shapes that move, resize, and rotate together as a unit in PowerPoint's
  UI. Child coordinate space defaults 1:1, so member shapes use ordinary
  slide-absolute EMU positions.
- **`Slide.Connect(from, fromSite, to, toSite, connectorType)`**
  (`p:cxnSp`) — connectors whose endpoints are *bound* to a shape's
  connection site (`stCxn`/`endCxn`), so dragging a connected shape in
  PowerPoint pulls the connector with it. Connection-site indices were
  extracted from a real python-pptx-generated file, not guessed.
  Currently supports `rect`/`roundRect`/`ellipse` endpoints — connecting
  to other preset geometries is rejected rather than silently mis-routed
  (see Known limitations below).

Caught in real PowerPoint: a connector's `a:xfrm` originally bounding-boxed
the two shapes' full rectangles, which drew a diagonal cutting through
both shapes on first paint (self-correcting only once a shape was
dragged). Fixed by spanning the two connection *points* instead.

---

## Known limitations / deferred work

- ⏸️ **Slide transitions** (`p:transition`) — deliberately deferred.
  Transitions only render during slideshow *playback*, not in any static
  capture (PDF export, thumbnail, edit view) this repo's verification
  pipeline can check — even opening the file in PowerPoint isn't enough
  without manually starting the show.
- ⏸️ **Rotation-aware connector endpoints** — `Slide.Connect`'s connection
  point currently reads a shape's *un-rotated* cardinal point. A connector
  to a rotated or flipped shape may render slightly detached on first
  paint (self-correcting once PowerPoint re-routes from the still-correct
  binding). Left undone rather than shipped as unverified rendering math —
  no rotated/connected shape has been checked in real PowerPoint yet.
- ⏸️ **Connector endpoints beyond `rect`/`roundRect`/`ellipse`** — other
  preset geometries (triangle, star, etc.) number their connection sites
  differently per the OOXML schema; binding to them needs the same
  extract-from-a-real-file verification the current three presets got
  before the guard can be widened.
- ⏸️ **Pattern fills** (`a:pattFill`) — greenfield, low value; explicitly
  left out of the completion batch's scope.
- ⏳ **Group nesting beyond shapes/text** — `Group` currently only routes
  `AddShape`/`AddTextBox` into a `p:grpSp`; nesting a picture, table, or
  sub-group needs its own real-PowerPoint check that the element actually
  moves with the group before it's added.
- ⏳ **Full-DOM reader** — `pptx.Open` supports navigation and text
  substitution; a complete object-model reader (beyond text) is unscoped.

## Explicitly out of scope

- **Placeholder inheritance resolution** — pptxgo is writer +
  template-substitution only. It emits correctly-typed `p:ph` elements;
  resolving inherited geometry/formatting is PowerPoint's job on open, by
  design (Phase 5).

---

*Generated 2026-07-21 from the actual package/test state of this repo,
cross-checked against session memory rather than copied from it.*
