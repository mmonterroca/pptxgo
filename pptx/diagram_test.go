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
	"bytes"
	"strconv"
	"strings"
	"testing"
)

// --- groups (p:grpSp) -------------------------------------------------------

func TestAddGroup_EmitsIdentityChildSpace(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddGroup(Inches(1), Inches(2), Inches(3), Inches(4))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, "<p:grpSp>") {
		t.Fatalf("expected p:grpSp, got %s", slide)
	}
	off := `<a:off x="` + strconv.Itoa(Inches(1)) + `" y="` + strconv.Itoa(Inches(2)) + `">`
	ext := `<a:ext cx="` + strconv.Itoa(Inches(3)) + `" cy="` + strconv.Itoa(Inches(4)) + `">`
	chOff := `<a:chOff x="` + strconv.Itoa(Inches(1)) + `" y="` + strconv.Itoa(Inches(2)) + `">`
	chExt := `<a:chExt cx="` + strconv.Itoa(Inches(3)) + `" cy="` + strconv.Itoa(Inches(4)) + `">`
	for _, want := range []string{off, ext, chOff, chExt} {
		if !strings.Contains(slide, want) {
			t.Errorf("expected %q (chOff=off, chExt=ext — the 1:1 identity child space), got %s", want, slide)
		}
	}
}

func TestGroup_AddShapeAppendsIntoGroupContentNotTopLevelSpTree(t *testing.T) {
	p := New()
	s := p.AddSlide()
	grp := s.AddGroup(Inches(1), Inches(1), Inches(4), Inches(2))
	grp.AddShape(ShapeRect, Inches(1), Inches(1), Inches(1), Inches(1)).
		AddParagraph().Text("Inside Group")
	s.AddShape(ShapeEllipse, Inches(6), Inches(1), Inches(1), Inches(1)).
		AddParagraph().Text("Outside Group")

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	grpIdx := strings.Index(slide, "<p:grpSp>")
	grpEndIdx := strings.Index(slide, "</p:grpSp>")
	insideIdx := strings.Index(slide, "Inside Group")
	outsideIdx := strings.Index(slide, "Outside Group")
	if grpIdx == -1 || grpEndIdx == -1 || insideIdx == -1 || outsideIdx == -1 {
		t.Fatalf("expected grpSp with both shapes' text present, got %s", slide)
	}
	if !(grpIdx < insideIdx && insideIdx < grpEndIdx) {
		t.Errorf("expected the group's own member shape nested inside <p:grpSp>...</p:grpSp>, got %s", slide)
	}
	if outsideIdx < grpEndIdx {
		t.Errorf("expected the top-level shape to come after the group closes, not nested inside it, got %s", slide)
	}
}

func TestGroup_MemberShapeUsesSlideAbsoluteCoordinates(t *testing.T) {
	// The group's 1:1 child space means a member shape's own (x, y) is the
	// same slide-absolute EMU position it would be outside the group — no
	// translation relative to the group's own origin.
	p := New()
	s := p.AddSlide()
	grp := s.AddGroup(Inches(1), Inches(1), Inches(4), Inches(2))
	grp.AddShape(ShapeRect, Inches(2), Inches(1.5), Inches(1), Inches(1))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:off x="`+strconv.Itoa(Inches(2))+`" y="`+strconv.Itoa(Inches(1.5))+`">`) {
		t.Errorf("expected the member shape's own off to be its literal slide-absolute (x, y), got %s", slide)
	}
}

func TestGroup_MembersShareSlideGlobalIDSequence(t *testing.T) {
	// Shape IDs are slide-global (Slide.allocID), not reset per group: the
	// root spTree's own nvGrpSpPr is a fixed id=1 (NewEmptySpTree, not part
	// of the slide's own counter), so allocation starts at 2 — the group
	// itself is 2, its two members are 3 and 4, and a shape added after the
	// group is 5, never colliding with a member's id.
	p := New()
	s := p.AddSlide()
	grp := s.AddGroup(Inches(1), Inches(1), Inches(4), Inches(2))
	grp.AddShape(ShapeRect, Inches(1), Inches(1), Inches(1), Inches(1))
	grp.AddShape(ShapeRect, Inches(2), Inches(1), Inches(1), Inches(1))
	s.AddShape(ShapeEllipse, Inches(6), Inches(1), Inches(1), Inches(1))

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	for _, want := range []string{
		`<p:cNvPr id="2" name="Group 2">`,
		`<p:cNvPr id="3" name="Shape 3">`,
		`<p:cNvPr id="4" name="Shape 4">`,
		`<p:cNvPr id="5" name="Shape 5">`,
	} {
		if !strings.Contains(slide, want) {
			t.Errorf("expected %q, got %s", want, slide)
		}
	}
}

func TestGroup_InvalidPresetGeometryAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	grp := s.AddGroup(Inches(1), Inches(1), Inches(4), Inches(2))
	grp.AddShape(PresetGeometry("bogus"), Inches(1), Inches(1), Inches(1), Inches(1))

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated invalid-preset error")
	}
}

// --- connectors (p:cxnSp) ---------------------------------------------------

func TestConnect_BindsStCxnAndEndCxnToShapeIDsAndSiteIndexes(t *testing.T) {
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(2), Inches(1))
	b := s.AddShape(ShapeRect, Inches(5), Inches(1), Inches(2), Inches(1))
	s.Connect(a, SiteRight, b, SiteLeft, ConnStraight)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, "<p:cxnSp>") {
		t.Fatalf("expected p:cxnSp, got %s", slide)
	}
	// a is shape id 2 (first shape after the root group's id=1), b is id 3.
	if !strings.Contains(slide, `<a:stCxn id="2" idx="3">`) {
		t.Errorf("expected stCxn bound to shape a's id at the right-site idx (3), got %s", slide)
	}
	if !strings.Contains(slide, `<a:endCxn id="3" idx="1">`) {
		t.Errorf("expected endCxn bound to shape b's id at the left-site idx (1), got %s", slide)
	}
	if !strings.Contains(slide, `<a:prstGeom prst="line">`) {
		t.Errorf("expected ConnStraight's prst=\"line\", got %s", slide)
	}
}

func TestConnect_BentConnectorEmitsBentPrstGeom(t *testing.T) {
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeEllipse, Inches(1), Inches(1), Inches(2), Inches(1))
	b := s.AddShape(ShapeEllipse, Inches(1), Inches(4), Inches(2), Inches(1))
	s.Connect(a, SiteBottom, b, SiteTop, ConnBent)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	if !strings.Contains(slide, `<a:prstGeom prst="bentConnector3">`) {
		t.Errorf("expected ConnBent's prst, got %s", slide)
	}
}

func TestConnect_XfrmSpansTheTwoConnectionPointsNotTheShapeBoundingBoxes(t *testing.T) {
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(1), Inches(1))
	b := s.AddShape(ShapeRect, Inches(5), Inches(3), Inches(1), Inches(1))
	s.Connect(a, SiteRight, b, SiteLeft, ConnStraight)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	cxnIdx := strings.Index(slide, "<p:cxnSp>")
	if cxnIdx == -1 {
		t.Fatalf("expected p:cxnSp, got %s", slide)
	}
	rest := slide[cxnIdx:]
	// a's right-center site is (Inches(2), Inches(1.5)); b's left-center
	// site is (Inches(5), Inches(3.5)) — the connector's own box spans
	// those two POINTS, not the two shapes' full rectangles (which would
	// start at Inches(1), not Inches(2)).
	wantOff := `<a:off x="` + strconv.Itoa(Inches(2)) + `" y="` + strconv.Itoa(Inches(1.5)) + `">`
	wantExt := `<a:ext cx="` + strconv.Itoa(Inches(3)) + `" cy="` + strconv.Itoa(Inches(2)) + `">`
	if !strings.Contains(rest, wantOff) {
		t.Errorf("expected off at the from-site point, got %s", rest)
	}
	if !strings.Contains(rest, wantExt) {
		t.Errorf("expected ext spanning from-site to to-site, got %s", rest)
	}
}

func TestConnect_AlignedShapesProduceAZeroHeightBoxNotADiagonal(t *testing.T) {
	// Regression: two same-height shapes at the same y, connected right-to-
	// left, must produce ext.cy=0 — a:prstGeom "line" draws the LITERAL
	// diagonal of whatever box it's given (no routing logic of its own), so
	// bounding-boxing the two shapes' full rectangles here (their heights
	// don't cancel out) drew a diagonal cutting through both shapes in real
	// PowerPoint's first paint, confirmed only by opening the file there —
	// see connectorXfrm's own doc comment.
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeRect, Inches(1), Inches(2), Inches(2), Inches(1))
	b := s.AddShape(ShapeRect, Inches(5), Inches(2), Inches(2), Inches(1))
	s.Connect(a, SiteRight, b, SiteLeft, ConnStraight)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	cxnIdx := strings.Index(slide, "<p:cxnSp>")
	if cxnIdx == -1 {
		t.Fatalf("expected p:cxnSp, got %s", slide)
	}
	if !strings.Contains(slide[cxnIdx:], `<a:ext cx="`+strconv.Itoa(Inches(2))+`" cy="0">`) {
		t.Errorf("expected a zero-height ext for two vertically-aligned connection points, got %s", slide[cxnIdx:])
	}
}

func TestConnect_ReversedDirectionSetsFlipHNotSwappedOff(t *testing.T) {
	// Connecting right-to-left (to is LEFT of from) must set FlipH rather
	// than silently producing an off/ext that reverses which shape the
	// connector's own begin/end point is bound to (ST_PositiveSize2D
	// forbids a negative Ext, so direction has to be carried by flipH/
	// flipV — see connectorXfrm's own doc comment, matching python-pptx's
	// own begin_x/end_x convention).
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeRect, Inches(5), Inches(1), Inches(1), Inches(1))
	b := s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(1), Inches(1))
	s.Connect(a, SiteLeft, b, SiteRight, ConnStraight)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	cxnIdx := strings.Index(slide, "<p:cxnSp>")
	if cxnIdx == -1 {
		t.Fatalf("expected p:cxnSp, got %s", slide)
	}
	if !strings.Contains(slide[cxnIdx:], `flipH="true"`) {
		t.Errorf("expected flipH=\"true\" for a right-to-left connection, got %s", slide[cxnIdx:])
	}
}

func TestConnect_InvalidSiteAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(1), Inches(1))
	b := s.AddShape(ShapeRect, Inches(5), Inches(1), Inches(1), Inches(1))
	s.Connect(a, ConnSite("center"), b, SiteLeft, ConnStraight)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated invalid-site error")
	}
}

func TestConnect_InvalidConnectorTypeAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(1), Inches(1))
	b := s.AddShape(ShapeRect, Inches(5), Inches(1), Inches(1), Inches(1))
	s.Connect(a, SiteRight, b, SiteLeft, ConnectorType("not-a-shape"))

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated invalid-connector-type error")
	}
}

func TestConnect_ShapeWithNoXfrmAccumulatesError(t *testing.T) {
	// A placeholder with inherited geometry has no a:xfrm of its own.
	p := New()
	s := p.AddSlide(WithLayout(LayoutTitleAndContent))
	title := s.AddPlaceholder(PlaceholderTitle, 0)
	b := s.AddShape(ShapeRect, Inches(5), Inches(1), Inches(1), Inches(1))
	s.Connect(title, SiteRight, b, SiteLeft, ConnStraight)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated no-xfrm error")
	}
}

func TestConnect_UnsupportedEndpointGeometryAccumulatesError(t *testing.T) {
	// A triangle numbers its cxnLst sites differently from rect/roundRect/
	// ellipse, so binding stCxn/endCxn with connSiteIdx's fixed cardinal
	// indices would silently point at the wrong site in real PowerPoint.
	// Connect rejects an endpoint drawn with any preset outside connSiteGeom
	// rather than emitting that schema-valid-but-mis-routed binding.
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(1), Inches(1))
	tri := s.AddShape(ShapeTriangle, Inches(5), Inches(1), Inches(1), Inches(1))
	s.Connect(a, SiteRight, tri, SiteLeft, ConnStraight)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated unsupported-geometry error")
	}
}

func TestConnect_RoundRectAndEllipseEndpointsAreAccepted(t *testing.T) {
	// The three verified presets (rect covered elsewhere) must NOT trip the
	// connSiteGeom guard — a roundRect-to-ellipse connection saves cleanly.
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeRoundRect, Inches(1), Inches(1), Inches(1), Inches(1))
	b := s.AddShape(ShapeEllipse, Inches(5), Inches(1), Inches(1), Inches(1))
	s.Connect(a, SiteRight, b, SiteLeft, ConnStraight)

	if err := p.Save(&bytes.Buffer{}); err != nil {
		t.Fatalf("expected roundRect/ellipse endpoints to be accepted, got %v", err)
	}
}

func TestConnectorRef_ReusesShapeRefLineStylingMethods(t *testing.T) {
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(1), Inches(1))
	b := s.AddShape(ShapeRect, Inches(5), Inches(1), Inches(1), Inches(1))
	s.Connect(a, SiteRight, b, SiteLeft, ConnStraight).
		Border(RGB(0, 0, 0), 1.5).
		BorderDash(DashDash).
		LineCap(LineCapRound).
		ArrowEnd(ArrowheadTriangle)

	files := generateFrom(t, p)
	slide := string(files["ppt/slides/slide1.xml"])

	cxnIdx := strings.Index(slide, "<p:cxnSp>")
	if cxnIdx == -1 {
		t.Fatalf("expected p:cxnSp, got %s", slide)
	}
	rest := slide[cxnIdx:]
	for _, want := range []string{`cap="rnd"`, `<a:prstDash val="dash">`, `type="triangle"`} {
		if !strings.Contains(rest, want) {
			t.Errorf("expected %q on the connector's own line, got %s", want, rest)
		}
	}
}

func TestConnectorRef_ArrowEndBeforeBorderAccumulatesError(t *testing.T) {
	p := New()
	s := p.AddSlide()
	a := s.AddShape(ShapeRect, Inches(1), Inches(1), Inches(1), Inches(1))
	b := s.AddShape(ShapeRect, Inches(5), Inches(1), Inches(1), Inches(1))
	s.Connect(a, SiteRight, b, SiteLeft, ConnStraight).ArrowEnd(ArrowheadTriangle)

	if err := p.Save(&bytes.Buffer{}); err == nil {
		t.Fatal("expected Save to return the accumulated no-outline error")
	}
}
