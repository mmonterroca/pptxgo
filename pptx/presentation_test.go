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
	"math"
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

func TestAddShape_UsesGivenPresetGeometryNotRect(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeEllipse, Inches(1), Inches(1), Inches(2), Inches(2)).
		AddParagraph().Text("On Track")

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:prstGeom prst="ellipse">`) {
		t.Errorf("expected prst=\"ellipse\", got %s", slide)
	}
	if strings.Contains(slide, `txBox="true"`) {
		t.Errorf("expected AddShape not to set the txBox marker (that's AddTextBox-only), got %s", slide)
	}
	if !strings.Contains(slide, "On Track") {
		t.Errorf("expected shape text content, got %s", slide)
	}
}

func TestAddTextBox_StillSetsTxBoxMarkerAndRectGeometry(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddTextBox(Inches(1), Inches(1), Inches(2), Inches(2))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:prstGeom prst="rect">`) {
		t.Errorf("expected prst=\"rect\", got %s", slide)
	}
	if !strings.Contains(slide, `txBox="true"`) {
		t.Errorf("expected the txBox marker preserved after the AddShape refactor, got %s", slide)
	}
}

func TestShapeRef_RotationConvertsDegreesTo60000ths(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).Rotation(45)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	// 45 degrees * 60,000 = 2,700,000.
	if !strings.Contains(slide, `rot="2700000"`) {
		t.Errorf("expected rot=\"2700000\" (45 degrees), got %s", slide)
	}
}

func TestShapeRef_NegativeRotationIsPreserved(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).Rotation(-90)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `rot="-5400000"`) {
		t.Errorf("expected rot=\"-5400000\" (-90 degrees), got %s", slide)
	}
}

func TestShapeRef_RotationNormalizesBeyondFullTurn(t *testing.T) {
	p := New()
	s := p.AddSlide()
	// 36000 degrees * 60,000 = 2,160,000,000, which overflows ST_Angle's
	// underlying 32-bit signed int (max 2,147,483,647) if emitted
	// un-normalized. 36000 mod 360 = 0, so this must normalize to rot="0"
	// (or simply omit the attribute, since 0 is Xfrm.Rot's zero value).
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).Rotation(36000)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if strings.Contains(slide, `rot="2160000000"`) {
		t.Fatalf("expected rotation to be normalized mod 360, not overflow ST_Angle, got %s", slide)
	}
	if err := p.Save(&bytes.Buffer{}); err != nil {
		t.Fatalf("expected no error for a large-but-finite rotation, got %v", err)
	}
}

func TestShapeRef_RotationBeyondOneTurnMatchesEquivalentAngle(t *testing.T) {
	p1 := New()
	s1 := p1.AddSlide()
	s1.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).Rotation(405)
	files1 := generateFrom(t, p1)

	p2 := New()
	s2 := p2.AddSlide()
	s2.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).Rotation(45)
	files2 := generateFrom(t, p2)

	// 405 degrees is the same rotation as 45 degrees (405 - 360 = 45).
	slide1 := string(files1["ppt/slides/slide1.xml"])
	slide2 := string(files2["ppt/slides/slide1.xml"])
	if slide1 != slide2 {
		t.Errorf("expected Rotation(405) to normalize identically to Rotation(45):\n405: %s\n45: %s", slide1, slide2)
	}
}

func TestShapeRef_RotationNaNAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).Rotation(math.NaN())

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated NaN-rotation error")
	}
}

func TestAddShape_InvalidPresetGeometryAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(PresetGeometry("rectangel"), Inches(1), Inches(1), Inches(2), Inches(2))

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated invalid-preset-geometry error")
	}
}

func TestAddShape_EveryDocumentedShapeConstantIsValid(t *testing.T) {
	for _, prst := range []PresetGeometry{
		ShapeRect, ShapeRoundRect, ShapeEllipse, ShapeTriangle, ShapeRightTriangle,
		ShapeParallelogram, ShapeTrapezoid, ShapeDiamond, ShapePentagon, ShapeHexagon,
		ShapeHeptagon, ShapeOctagon, ShapeStar4, ShapeStar5, ShapeStar6, ShapeStar8,
		ShapeRightArrow, ShapeLeftArrow, ShapeUpArrow, ShapeDownArrow, ShapeLeftRightArrow,
		ShapeUpDownArrow, ShapeChevron, ShapeDonut, ShapeNoSmoking, ShapeHeart,
		ShapeLightningBolt, ShapeSun, ShapeMoon, ShapeCloud, ShapeArc, ShapePlaque,
		ShapeCan, ShapeCube, ShapeBevel, ShapeSmileyFace, ShapeWave, ShapeDoubleWave,
	} {
		if !IsValidPresetGeometry(prst) {
			t.Errorf("expected %q (a documented Shape* constant) to be a valid ST_ShapeType", prst)
		}
	}
}

func TestShapeRef_FlipHAndFlipVSetXfrmAttrs(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).FlipH().FlipV()

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `flipH="true"`) {
		t.Errorf("expected flipH=\"true\", got %s", slide)
	}
	if !strings.Contains(slide, `flipV="true"`) {
		t.Errorf("expected flipV=\"true\", got %s", slide)
	}
}

func TestParagraph_BulletEmitsCharAndFontBeforeChar(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("First item").Bullet("•", "Arial").Indent(18, -18).Level(1)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:buChar char="•">`) && !strings.Contains(slide, `char="•"`) {
		t.Errorf("expected buChar with the bullet glyph, got %s", slide)
	}
	if !strings.Contains(slide, `typeface="Arial"`) {
		t.Errorf("expected buFont typeface=\"Arial\", got %s", slide)
	}
	buFontIdx := strings.Index(slide, "<a:buFont")
	buCharIdx := strings.Index(slide, "<a:buChar")
	if buFontIdx == -1 || buCharIdx == -1 || buFontIdx > buCharIdx {
		t.Errorf("expected a:buFont before a:buChar, got %s", slide)
	}
	if !strings.Contains(slide, `lvl="1"`) {
		t.Errorf("expected lvl=\"1\", got %s", slide)
	}
}

func TestParagraph_NumberedBulletSetsAutoNumType(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("Step one").NumberedBullet(NumArabicPeriod)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:buAutoNum type="arabicPeriod">`) {
		t.Errorf("expected buAutoNum type=\"arabicPeriod\", got %s", slide)
	}
}

func TestParagraph_NoBulletOverridesEarlierBulletCall(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("Plain").Bullet("•", "Arial").NoBullet()

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if strings.Contains(slide, "a:buChar") || strings.Contains(slide, "a:buFont") {
		t.Errorf("expected NoBullet to clear the earlier Bullet call, got %s", slide)
	}
	if !strings.Contains(slide, "<a:buNone>") {
		t.Errorf("expected a:buNone, got %s", slide)
	}
}

func TestParagraph_SpacingEmitsPercentAndPoints(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("Spaced").LineSpacing(150).SpaceBefore(6).SpaceAfter(12)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:spcPct val="150000">`) {
		t.Errorf("expected lnSpc spcPct val=\"150000\" (150%%), got %s", slide)
	}
	// 6pt and 12pt in hundredths of a point.
	if !strings.Contains(slide, `<a:spcPts val="600">`) {
		t.Errorf("expected spcBef spcPts val=\"600\" (6pt), got %s", slide)
	}
	if !strings.Contains(slide, `<a:spcPts val="1200">`) {
		t.Errorf("expected spcAft spcPts val=\"1200\" (12pt), got %s", slide)
	}
}

func TestShapeRef_AutofitSwitchesAmongTheThreeModes(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2)).Autofit(AutofitShrinkText)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, "<a:normAutofit>") {
		t.Errorf("expected a:normAutofit, got %s", slide)
	}
	if strings.Contains(slide, "a:noAutofit") || strings.Contains(slide, "a:spAutoFit") {
		t.Errorf("expected only normAutofit present, got %s", slide)
	}
}

func TestShapeRef_InsetsAnchorAndWordWrap(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2)).
		Insets(10, 5, 10, 5).
		Anchor(AnchorMiddle).
		WordWrap(false)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `anchor="ctr"`) {
		t.Errorf("expected anchor=\"ctr\", got %s", slide)
	}
	if !strings.Contains(slide, `wrap="none"`) {
		t.Errorf("expected wrap=\"none\", got %s", slide)
	}
	// 10pt = 127000 EMU, 5pt = 63500 EMU.
	if !strings.Contains(slide, `lIns="127000"`) || !strings.Contains(slide, `tIns="63500"`) {
		t.Errorf("expected lIns=\"127000\" and tIns=\"63500\", got %s", slide)
	}
}

func TestShapeRef_FillSchemeEmitsSchemeClrNotSrgbClr(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		FillScheme(SchemeAccent1).
		BorderScheme(SchemeDark2, 1.0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:schemeClr val="accent1">`) {
		t.Errorf("expected fill schemeClr val=\"accent1\", got %s", slide)
	}
	if !strings.Contains(slide, `<a:schemeClr val="dk2">`) {
		t.Errorf("expected border schemeClr val=\"dk2\", got %s", slide)
	}
	if strings.Contains(slide, "a:srgbClr") {
		t.Errorf("expected no a:srgbClr when using scheme colors, got %s", slide)
	}
}

func TestShapeRef_FillThenNoFillClearsSolidFill(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		Fill(RGB(0xFF, 0, 0)).
		NoFill()

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, "<a:noFill>") {
		t.Errorf("expected a:noFill, got %s", slide)
	}
	if strings.Contains(slide, "a:solidFill") {
		t.Errorf("expected NoFill to clear the earlier Fill call, got %s", slide)
	}
}

func TestShapeRef_NoFillThenFillClearsNoFill(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).
		NoFill().
		Fill(RGB(0, 0xFF, 0))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if strings.Contains(slide, "a:noFill") {
		t.Errorf("expected Fill to clear the earlier NoFill call, got %s", slide)
	}
	if !strings.Contains(slide, "a:solidFill") {
		t.Errorf("expected a:solidFill present, got %s", slide)
	}
}

func TestParagraph_ColorSchemeEmitsSchemeClr(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("Themed").ColorScheme(SchemeAccent2)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:schemeClr val="accent2">`) {
		t.Errorf("expected run color schemeClr val=\"accent2\", got %s", slide)
	}
}

func TestSlide_BackgroundEmitsBgBeforeSpTree(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.Background(RGB(0x1F, 0x49, 0x7D))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	bgIdx := strings.Index(slide, "<p:bg>")
	spTreeIdx := strings.Index(slide, "<p:spTree>")
	if bgIdx == -1 || spTreeIdx == -1 || bgIdx > spTreeIdx {
		t.Errorf("expected p:bg before p:spTree, got %s", slide)
	}
	if !strings.Contains(slide, `<a:srgbClr val="1F497D">`) {
		t.Errorf("expected background srgbClr val=\"1F497D\", got %s", slide)
	}
}

func TestSlide_BackgroundSchemeEmitsSchemeClr(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.BackgroundScheme(SchemeLight1)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:schemeClr val="lt1">`) {
		t.Errorf("expected background schemeClr val=\"lt1\", got %s", slide)
	}
}

func TestSlide_NoBackgroundOmitsBgElement(t *testing.T) {
	p := New()
	p.AddSlide()

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if strings.Contains(slide, "<p:bg>") {
		t.Errorf("expected no p:bg when Background is never called, got %s", slide)
	}
}

func TestShapeRef_InsetsExplicitZeroIsNotDroppedByOmitempty(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2)).Insets(0, 0, 0, 0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	for _, want := range []string{`lIns="0"`, `tIns="0"`, `rIns="0"`, `bIns="0"`} {
		if !strings.Contains(slide, want) {
			t.Errorf("expected explicit zero inset %s to marshal (not be dropped as unset), got %s", want, slide)
		}
	}
}

func TestShapeRef_InsetsRoundsRatherThanTruncates(t *testing.T) {
	p := New()
	s := p.AddSlide()
	// 2.3pt * 12700 EMU/pt = 29209.999999999996 in float64 — direct int()
	// truncates to 29209, but the correct rounded EMU value is 29210.
	s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2)).Insets(2.3, 0, 0, 0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `lIns="29210"`) {
		t.Errorf("expected lIns=\"29210\" (rounded, not truncated to 29209), got %s", slide)
	}
}

func TestParagraph_IndentExplicitZeroIsNotDroppedByOmitempty(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("no indent").Indent(0, 0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `marL="0"`) || !strings.Contains(slide, `indent="0"`) {
		t.Errorf("expected marL=\"0\" and indent=\"0\" to marshal (not be dropped as unset), got %s", slide)
	}
}

func TestParagraph_IndentRoundsRatherThanTruncates(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("x").Indent(2.3, 0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `marL="29210"`) {
		t.Errorf("expected marL=\"29210\" (rounded, not truncated to 29209), got %s", slide)
	}
}

func TestParagraph_LevelOutOfRangeAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("x").Level(9)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated out-of-range Level error")
	}
}

func TestParagraph_LevelNegativeAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("x").Level(-1)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated negative-Level error")
	}
}

func TestParagraph_LevelBoundaryValuesDoNotError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("a").Level(0)
	tb.AddParagraph().Text("b").Level(8)

	if err := p.Save(&bytes.Buffer{}); err != nil {
		t.Fatalf("expected no error for boundary-valid levels 0 and 8, got %v", err)
	}
}

func TestParagraph_LevelZeroMarshalsExplicitly(t *testing.T) {
	p := New()
	s := p.AddSlide()
	tb := s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2))
	tb.AddParagraph().Text("x").Level(0)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `lvl="0"`) {
		t.Errorf("expected an explicit Level(0) call to marshal lvl=\"0\" (not be dropped by omitempty), got %s", slide)
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

func TestBorder_NaNWidthAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	// NaN fails both the "< 0" and "> max" comparisons (every comparison
	// against NaN is false), so it must be checked explicitly rather than
	// falling through to the range check.
	s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2)).Border(RGB(0, 0, 0), math.NaN())

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated NaN-width error")
	}
}

func TestBorderScheme_NaNWidthAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).BorderScheme(SchemeAccent1, math.NaN())

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated NaN-width error")
	}
}

func TestBorder_RoundsRatherThanTruncates(t *testing.T) {
	p := New()
	s := p.AddSlide()
	// 2.3pt * 12700 EMU/pt = 29209.999999999996 in float64 — direct int()
	// truncates to 29209, but the correct rounded EMU value is 29210.
	s.AddTextBox(Inches(1), Inches(1), Inches(8), Inches(2)).Border(RGB(0, 0, 0), 2.3)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:ln w="29210">`) {
		t.Errorf("expected a:ln w=\"29210\" (rounded, not truncated to 29209), got %s", slide)
	}
}

func TestSchemeColor_Bg1Tx1Bg2Tx2AreDistinctFromDk1Lt1(t *testing.T) {
	// bg1/tx1/bg2/tx2 are aliases (through the default color map) for
	// lt1/dk1/lt2/dk2, not new theme slots — but they are distinct valid
	// ST_SchemeColorVal strings, and must marshal as given rather than
	// being silently rewritten to their dk/lt equivalent.
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(2)).FillScheme(SchemeBackground1)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:schemeClr val="bg1">`) {
		t.Errorf("expected schemeClr val=\"bg1\", got %s", slide)
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
