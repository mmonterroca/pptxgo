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
	"archive/zip"
	"bytes"
	"encoding/xml"
	"image/png"
	"os"
	"regexp"
	"strings"
	"testing"
)

// This is the Go-level regression suite. It cannot substitute for the
// OpenXML SDK validator or an actual OOXML consumer (see PptxValidator/ and
// the manual open-in-PowerPoint step) — a document can satisfy every check
// below and still be schema-invalid. What it catches cheaply and on every
// `go test`, without any external tool, is exactly the class of bug that
// caused docxgo's own round-trip incidents: a relationship pointing at a
// part that doesn't exist, or a part nobody points at.

func generate(t *testing.T) map[string][]byte {
	t.Helper()
	p := New()
	p.AddSlide()
	return generateFrom(t, p)
}

// generateFrom saves p and unzips the result, for tests that need a
// presentation shaped differently than generate's single blank slide.
func generateFrom(t *testing.T, p *Presentation) map[string][]byte {
	t.Helper()
	var buf bytes.Buffer
	if err := p.Save(&buf); err != nil {
		t.Fatalf("Save: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("zip.NewReader: %v", err)
	}

	files := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("open %s: %v", f.Name, err)
		}
		var b bytes.Buffer
		if _, err := b.ReadFrom(rc); err != nil {
			t.Fatalf("read %s: %v", f.Name, err)
		}
		rc.Close()
		files[f.Name] = b.Bytes()
	}
	return files
}

func TestNew_ProducesEveryRequiredPart(t *testing.T) {
	files := generate(t)

	for _, want := range []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels",
		"ppt/theme/theme1.xml",
		"ppt/slideMasters/slideMaster1.xml",
		"ppt/slideMasters/_rels/slideMaster1.xml.rels",
		"ppt/slideLayouts/slideLayout1.xml",
		"ppt/slideLayouts/_rels/slideLayout1.xml.rels",
		"ppt/slides/slide1.xml",
		"ppt/slides/_rels/slide1.xml.rels",
		"docProps/core.xml",
		"docProps/app.xml",
	} {
		if _, ok := files[want]; !ok {
			t.Errorf("missing required part %s", want)
		}
	}
}

func TestNew_EveryPartIsWellFormedXML(t *testing.T) {
	for name, data := range generate(t) {
		if !strings.HasSuffix(name, ".xml") && !strings.HasSuffix(name, ".rels") {
			continue
		}
		var v any
		if err := xml.Unmarshal(data, &v); err != nil {
			t.Errorf("%s is not well-formed XML: %v", name, err)
		}
	}
}

var relTargetRe = regexp.MustCompile(`Target="([^"]+)"\s*(?:TargetMode="([^"]+)"\s*)?/?>`)

// resolveRelTarget resolves a relationship's Target attribute against the
// part that owns the .rels file it came from, following the OPC convention
// (relative to the owning part's directory, not the package root).
func resolveRelTarget(relsPath, target string) string {
	// relsPath is ".../_rels/<owner-basename>.rels" or "_rels/.rels".
	dir := strings.TrimSuffix(relsPath, "/_rels/"+lastSegment(relsPath))
	if dir == relsPath { // root: "_rels/.rels"
		dir = ""
	}
	segments := strings.Split(dir, "/")
	if dir == "" {
		segments = nil
	}
	for _, part := range strings.Split(target, "/") {
		switch part {
		case ".":
			// no-op
		case "..":
			if len(segments) > 0 {
				segments = segments[:len(segments)-1]
			}
		default:
			segments = append(segments, part)
		}
	}
	return strings.Join(segments, "/")
}

func lastSegment(p string) string {
	i := strings.LastIndex(p, "/")
	return p[i+1:]
}

func TestNew_EveryInternalRelationshipTargetExists(t *testing.T) {
	files := generate(t)

	for name, data := range files {
		if !strings.HasSuffix(name, ".rels") {
			continue
		}
		for _, m := range relTargetRe.FindAllStringSubmatch(string(data), -1) {
			target, mode := m[1], m[2]
			if mode == "External" {
				continue
			}
			resolved := resolveRelTarget(name, target)
			if _, ok := files[resolved]; !ok {
				t.Errorf("%s references target %q, resolved to %q, which does not exist as a part",
					name, target, resolved)
			}
		}
	}
}

func TestNew_ContentTypesCoversEveryNonMediaPart(t *testing.T) {
	files := generate(t)
	ct := string(files["[Content_Types].xml"])

	for name := range files {
		if strings.HasSuffix(name, ".rels") || name == "[Content_Types].xml" {
			continue
		}
		if !strings.Contains(ct, `PartName="/`+name+`"`) {
			t.Errorf("[Content_Types].xml has no Override for part %s", name)
		}
	}
}

func TestNew_RelationshipIDsAreUniquePerOwner(t *testing.T) {
	idRe := regexp.MustCompile(`Id="([^"]+)"`)
	for name, data := range generate(t) {
		if !strings.HasSuffix(name, ".rels") {
			continue
		}
		seen := make(map[string]bool)
		for _, m := range idRe.FindAllStringSubmatch(string(data), -1) {
			id := m[1]
			if seen[id] {
				t.Errorf("%s has a duplicate relationship ID %s", name, id)
			}
			seen[id] = true
		}
	}
}

func TestNew_HasNoSlideUntilAddSlideIsCalled(t *testing.T) {
	files := generateFrom(t, New())

	if _, ok := files["ppt/slides/slide1.xml"]; ok {
		t.Error("expected no slide part before AddSlide is called")
	}
	if strings.Contains(string(files["ppt/presentation.xml"]), "p:sldIdLst") {
		t.Error("expected no p:sldIdLst in presentation.xml before AddSlide is called")
	}
}

func TestAddSlide_AssignsSequentialIDsAndRelationships(t *testing.T) {
	p := New()
	p.AddSlide()
	p.AddSlide()
	files := generateFrom(t, p)

	for _, want := range []string{"ppt/slides/slide1.xml", "ppt/slides/slide2.xml"} {
		if _, ok := files[want]; !ok {
			t.Errorf("missing %s", want)
		}
	}

	pres := string(files["ppt/presentation.xml"])
	for _, want := range []string{`id="256"`, `id="257"`} {
		if !strings.Contains(pres, want) {
			t.Errorf("expected sldId %s in presentation.xml, got %s", want, pres)
		}
	}

	rels := string(files["ppt/_rels/presentation.xml.rels"])
	for _, want := range []string{`Target="slides/slide1.xml"`, `Target="slides/slide2.xml"`} {
		if !strings.Contains(rels, want) {
			t.Errorf("expected %s in presentation.xml.rels, got %s", want, rels)
		}
	}
}

func TestAddTextBox_EmitsSchemaOrderedFormattedText(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().
		Text("Quarterly Results").Bold().FontSize(32).Font("Calibri").Color(RGB(0x1F, 0x49, 0x7D)).
		Alignment(AlignCenter)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	for _, want := range []string{
		"<p:sp>",
		"<p:txBody>",
		`<a:t xml:space="preserve">Quarterly Results</a:t>`,
		`sz="3200"`,
		`b="1"`,
		`algn="ctr"`,
		`typeface="Calibri"`,
	} {
		if !strings.Contains(slide, want) {
			t.Errorf("expected %q in slide1.xml, got %s", want, slide)
		}
	}

	fillIdx := strings.Index(slide, "<a:solidFill")
	latinIdx := strings.Index(slide, "<a:latin")
	if fillIdx == -1 || latinIdx == -1 || fillIdx > latinIdx {
		t.Errorf("expected a:solidFill before a:latin, got %s", slide)
	}

	if !strings.Contains(slide, `id="2"`) {
		t.Errorf("expected the text box's shape id to start at 2, got %s", slide)
	}
}

func TestAddTextBox_WithNoParagraphsStillEmitsOneEmptyParagraph(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, "<a:p></a:p>") && !strings.Contains(slide, "<a:p/>") {
		t.Errorf("expected a fallback empty paragraph, got %s", slide)
	}
}

func TestAddImageFromBytes_EmbedsMediaAndWiresRelationship(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddImageFromBytes(pngBytes(t, 120, 80), Inches(1), Inches(1))

	files := generateFrom(t, p)

	if _, ok := files["ppt/media/image1.png"]; !ok {
		t.Fatalf("expected ppt/media/image1.png to exist, got %v", mapKeys(files))
	}

	slide := string(files["ppt/slides/slide1.xml"])
	for _, want := range []string{"<p:pic>", "<p:blipFill>", `r:embed="rId2"`} {
		if !strings.Contains(slide, want) {
			t.Errorf("expected %q in slide1.xml, got %s", want, slide)
		}
	}

	// 120px/80px at 96 DPI -> 1143000/762000 EMU.
	if !strings.Contains(slide, `cx="1143000" cy="762000"`) {
		t.Errorf("expected auto-computed 96 DPI size, got %s", slide)
	}

	ct := string(files["[Content_Types].xml"])
	if !strings.Contains(ct, `Extension="png"`) {
		t.Errorf("expected a png Default in [Content_Types].xml, got %s", ct)
	}

	rels := string(files["ppt/slides/_rels/slide1.xml.rels"])
	if !strings.Contains(rels, `Target="../media/image1.png"`) {
		t.Errorf("expected image relationship target, got %s", rels)
	}
}

func TestAddImageFromBytesWithSize_UsesExplicitDimensions(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddImageFromBytesWithSize(pngBytes(t, 120, 80), Inches(1), Inches(1), Inches(3), Inches(2))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])
	if !strings.Contains(slide, `cx="2743200" cy="1828800"`) {
		t.Errorf("expected explicit 3in x 2in size, got %s", slide)
	}
}

func TestPictureRef_Border_EmitsLnAfterPrstGeom(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddImageFromBytes(pngBytes(t, 40, 40), Inches(1), Inches(1)).
		Border(RGB(0, 0, 0), 1.0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	prstGeomIdx := strings.Index(slide, "<a:prstGeom")
	lnIdx := strings.Index(slide, "<a:ln")
	if prstGeomIdx == -1 || lnIdx == -1 || prstGeomIdx > lnIdx {
		t.Errorf("expected a:prstGeom before a:ln, got %s", slide)
	}
	// 1 point = 12700 EMU.
	if !strings.Contains(slide, `<a:ln w="12700">`) {
		t.Errorf("expected a 1pt (12700 EMU) border width, got %s", slide)
	}
}

func TestTextBox_FillAndBorder_EmitsSolidFillBeforeLn(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2)).
		Fill(RGB(0xE7, 0xE6, 0xE6)).
		Border(RGB(0x1F, 0x49, 0x7D), 1.5)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	fillIdx := strings.Index(slide, `<a:srgbClr val="E7E6E6">`)
	lnIdx := strings.Index(slide, "<a:ln")
	if fillIdx == -1 || lnIdx == -1 || fillIdx > lnIdx {
		t.Errorf("expected shape fill before a:ln, got %s", slide)
	}
	// 1.5 points = 19050 EMU.
	if !strings.Contains(slide, `<a:ln w="19050">`) {
		t.Errorf("expected a 1.5pt (19050 EMU) border width, got %s", slide)
	}
}

func TestAddImage_MissingFileAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddImage("/nonexistent/path/does-not-exist.png", Inches(1), Inches(1))

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated file-read error")
	}
}

func TestAddImageFromBytes_NonImageDataAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddImageFromBytes([]byte("not an image"), Inches(1), Inches(1))

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated format-detection error")
	}
}

func TestAddImageFromBytes_NilOrEmptyDataAccumulatesError(t *testing.T) {
	// A nil/empty byte slice must record an error, not silently emit a
	// p:pic with no p:blipFill (which CT_Picture forbids) that Save would
	// then happily write as a schema-invalid file.
	for _, tc := range []struct {
		name string
		data []byte
	}{
		{"nil", nil},
		{"empty", []byte{}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			p := New()
			s := p.AddSlide()
			s.AddImageFromBytes(tc.data, Inches(1), Inches(1))
			if err := p.Save(&bytes.Buffer{}); err == nil {
				t.Fatalf("expected Save to error for %s image data", tc.name)
			}
		})
	}

	// The WithSize variant takes the same guard.
	p := New()
	s := p.AddSlide()
	s.AddImageFromBytesWithSize(nil, Inches(1), Inches(1), Inches(2), Inches(2))
	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to error for nil image data (WithSize)")
	}
}

func TestAddImageFromBytesWithSize_ExplicitZeroSizeIsPreserved(t *testing.T) {
	p := New()
	s := p.AddSlide()
	// An explicit (0, 0) is a caller choice, distinct from "no size given"
	// (AddImageFromBytes) — it must not silently fall back to the image's
	// own 96 DPI dimensions.
	s.AddImageFromBytesWithSize(pngBytes(t, 120, 80), Inches(1), Inches(1), 0, 0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])
	if !strings.Contains(slide, `cx="0" cy="0"`) {
		t.Errorf("expected the explicit 0x0 size to be preserved, got %s", slide)
	}
}

func TestAddImageWithSize_ExplicitZeroSizeIsPreserved(t *testing.T) {
	p := New()
	s := p.AddSlide()
	f, err := os.CreateTemp(t.TempDir(), "*.png")
	if err != nil {
		t.Fatalf("CreateTemp: %v", err)
	}
	if _, err := f.Write(pngBytes(t, 120, 80)); err != nil {
		t.Fatalf("Write: %v", err)
	}
	f.Close()

	s.AddImageWithSize(f.Name(), Inches(1), Inches(1), 0, 0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])
	if !strings.Contains(slide, `cx="0" cy="0"`) {
		t.Errorf("expected the explicit 0x0 size to be preserved, got %s", slide)
	}
}

func TestBorder_NegativeWidthAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2)).Border(RGB(0, 0, 0), -1)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated border-width error")
	}
}

func TestBorder_OverMaxWidthAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddImageFromBytes(pngBytes(t, 40, 40), Inches(1), Inches(1)).Border(RGB(0, 0, 0), 1584.1)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated border-width error")
	}
}

func TestBorder_BoundaryWidthsDoNotError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2)).Border(RGB(0, 0, 0), 0)
	s.AddTextBox(Inches(1), Inches(4), Inches(8), Inches(2)).Border(RGB(0, 0, 0), 1584)

	if err := p.Save(&bytes.Buffer{}); err != nil {
		t.Fatalf("expected no error for boundary-valid border widths, got %v", err)
	}
}

func TestAddTable_EmitsSchemaOrderedGraphicFrame(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tbl := s.AddTable(2, 3, Inches(1), Inches(1), Inches(9), Inches(2))
	tbl.Cell(0, 0).Text("Q1")
	tbl.Cell(0, 1).Text("Q2")
	tbl.Cell(1, 0).Text("100")

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	nvIdx := strings.Index(slide, "<p:nvGraphicFramePr>")
	xfrmIdx := strings.Index(slide, "<p:xfrm>")
	graphicIdx := strings.Index(slide, "<a:graphic")
	tblIdx := strings.Index(slide, "<a:tbl>")
	tblGridIdx := strings.Index(slide, "<a:tblGrid>")
	if nvIdx == -1 || xfrmIdx == -1 || graphicIdx == -1 || tblIdx == -1 || tblGridIdx == -1 {
		t.Fatalf("expected nvGraphicFramePr, xfrm, graphic, tbl, and tblGrid all present, got %s", slide)
	}
	if !(nvIdx < xfrmIdx && xfrmIdx < graphicIdx && graphicIdx < tblIdx && tblIdx < tblGridIdx) {
		t.Errorf("expected nvGraphicFramePr < xfrm < graphic < tbl < tblGrid order, got %s", slide)
	}
	if !strings.Contains(slide, `uri="http://schemas.openxmlformats.org/drawingml/2006/table"`) {
		t.Errorf("expected the table graphicData URI, got %s", slide)
	}
	for _, want := range []string{"Q1", "Q2", "100"} {
		if !strings.Contains(slide, want) {
			t.Errorf("expected cell text %q, got %s", want, slide)
		}
	}
	// 3 columns splitting 9in (8229600 EMU) evenly -> 2743200 EMU each.
	if !strings.Contains(slide, `<a:gridCol w="2743200">`) {
		t.Errorf("expected evenly-split column width 2743200 EMU, got %s", slide)
	}
}

func TestTable_ColumnWidthAndRowHeightOverrideEvenSplit(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tbl := s.AddTable(2, 2, Inches(1), Inches(1), Inches(4), Inches(2))
	tbl.ColumnWidth(0, Inches(3))
	tbl.RowHeight(0, Inches(1.5))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:gridCol w="2743200">`) { // 3in
		t.Errorf("expected overridden column 0 width 2743200 EMU (3in), got %s", slide)
	}
	if !strings.Contains(slide, `<a:tr h="1371600">`) { // 1.5in
		t.Errorf("expected overridden row 0 height 1371600 EMU (1.5in), got %s", slide)
	}
}

func TestAddTable_DefaultsToFirstRowBandedStyle(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddTable(1, 1, Inches(1), Inches(1), Inches(2), Inches(1))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `firstRow="1"`) || !strings.Contains(slide, `bandRow="1"`) {
		t.Errorf("expected firstRow and bandRow both set, got %s", slide)
	}
	if !strings.Contains(slide, "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}") {
		t.Errorf("expected the default table style GUID, got %s", slide)
	}
}

// pngBytes returns a solid-color w x h PNG, encoded in memory.
func pngBytes(t *testing.T, w, h int) []byte {
	t.Helper()
	var buf bytes.Buffer
	if err := png.Encode(&buf, solidImage(w, h)); err != nil {
		t.Fatalf("png.Encode: %v", err)
	}
	return buf.Bytes()
}

func mapKeys(m map[string][]byte) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	return keys
}
