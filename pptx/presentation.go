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
	"crypto/sha256"
	"io"
	"strconv"

	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/opc"
)

// Slide 16:9 widescreen canvas size (13.333in x 7.5in) and notes page size
// (7.5in x 10in), in EMUs — PowerPoint's own modern defaults.
const (
	slideWidthEMU  = 12192000
	slideHeightEMU = 6858000
	notesWidthEMU  = 6858000
	notesHeightEMU = 9144000
)

// Standard slide sizes, in EMUs, for use with WithSlideSize.
const (
	SlideSizeWidescreen16x9Width  = slideWidthEMU // 13.333in x 7.5in — PowerPoint's modern default (New's own default)
	SlideSizeWidescreen16x9Height = slideHeightEMU
	SlideSizeStandard4x3Width     = 9144000 // 10in x 7.5in — the pre-2013 PowerPoint default
	SlideSizeStandard4x3Height    = 6858000
)

// Conventional ID ranges: PowerPoint reserves the low range of the 32-bit
// ID space and starts slide master IDs at 2^31. Slide IDs and layout IDs
// need only be unique within the presentation; these starting points match
// what every real Office-produced file uses.
const (
	firstSldMasterID = 2147483648
	firstSldLayoutID = 2147483649
	firstSldID       = 256
)

// Presentation is a PresentationML document under construction: one theme,
// one slide master, and one slide layout — the structural backbone every
// presentation needs regardless of slide count — plus whatever slides
// AddSlide adds. New() starts with zero slides.
type Presentation struct {
	pkg         *opc.Package
	pres        *XMLPresentation
	presRels    *opc.RelationshipManager
	slideCount  int
	errs        []error
	mediaByHash map[[sha256.Size]byte]string // content hash -> already-embedded media part's ppt/media/ basename; see mediaBasename
}

// Option configures a Presentation at construction time, for use with New.
type Option func(*presentationConfig)

// presentationConfig collects Option values before New builds anything —
// keeping every option's effect in one place rather than threading a
// growing parameter list through New's body.
type presentationConfig struct {
	slideWidthEMU  int
	slideHeightEMU int
	slideSizeType  string // ST_SlideSizeType, e.g. "screen16x9"; "" when the size doesn't match a named preset
}

// WithSlideSize overrides New's default 16:9 widescreen canvas (13.333in x
// 7.5in) with an explicit width and height, in EMUs (see the Inches
// helper) — including a portrait layout, by simply passing a height
// greater than the width. The resulting p:sldSz carries no type
// attribute, since ST_SlideSizeType names a fixed set of standard sizes
// and an arbitrary custom size doesn't correspond to any of them.
func WithSlideSize(widthEMU, heightEMU int) Option {
	return func(c *presentationConfig) {
		c.slideWidthEMU = widthEMU
		c.slideHeightEMU = heightEMU
		c.slideSizeType = ""
	}
}

// WithStandard4x3 sets the slide canvas to the pre-2013 PowerPoint default,
// 10in x 7.5in (4:3), instead of New's own 16:9 widescreen default.
func WithStandard4x3() Option {
	return func(c *presentationConfig) {
		c.slideWidthEMU = SlideSizeStandard4x3Width
		c.slideHeightEMU = SlideSizeStandard4x3Height
		c.slideSizeType = "screen4x3"
	}
}

// New builds a presentation with its theme, slide master, and slide layout
// already wired, and no slides. Call AddSlide to add content. With no
// options, the slide canvas is 16:9 widescreen; pass WithSlideSize or
// WithStandard4x3 to override it.
func New(opts ...Option) *Presentation {
	cfg := &presentationConfig{
		slideWidthEMU:  slideWidthEMU,
		slideHeightEMU: slideHeightEMU,
		slideSizeType:  "screen16x9",
	}
	for _, opt := range opts {
		opt(cfg)
	}

	pkg := opc.NewPackage()

	pkg.AddRawPart(PathTheme1, opc.ContentTypeTheme, []byte(defaultTheme))

	// Add the master's relationships first and thread the generated rIds
	// into the XML, rather than restating "rId2" as a literal that silently
	// depends on the exact order of these Add calls. The master body only
	// references the layout by rId; the theme rel lives in the .rels but is
	// not referenced from the body.
	masterRels := pkg.Relationships(PathSlideMaster1)
	if _, err := masterRels.Add(opc.RelTypeTheme, "../theme/theme1.xml", "Internal"); err != nil {
		panic(err) // static, well-formed arguments; cannot fail
	}
	layoutRID, err := masterRels.Add(RelTypeSlideLayout, "../slideLayouts/slideLayout1.xml", "Internal")
	if err != nil {
		panic(err)
	}
	pkg.AddPart(PathSlideMaster1, ContentTypeSlideMaster, &XMLSlideMaster{
		XmlnsA: drawingml.NamespaceMain,
		XmlnsR: drawingml.NamespaceRelationships,
		XmlnsP: NamespaceMain,
		CSld:   &CSld{SpTree: NewEmptySpTree()},
		ClrMap: NewDefaultClrMap(),
		SldLayoutIdLst: &SldLayoutIdLst{
			Entries: []*SldLayoutId{{ID: firstSldLayoutID, RID: layoutRID}},
		},
		TxStyles: NewDefaultTxStyles(),
	})

	pkg.AddPart(PathSlideLayout1, ContentTypeSlideLayout, &XMLSlideLayout{
		XmlnsA:    drawingml.NamespaceMain,
		XmlnsR:    drawingml.NamespaceRelationships,
		XmlnsP:    NamespaceMain,
		Type:      "blank",
		CSld:      &CSld{SpTree: NewEmptySpTree()},
		ClrMapOvr: NewClrMapOvrInherit(),
	})
	if _, err := pkg.Relationships(PathSlideLayout1).Add(RelTypeSlideMaster, "../slideMasters/slideMaster1.xml", "Internal"); err != nil {
		panic(err)
	}

	// Same pattern: add the presentation's relationships and reference the
	// master by the rId Add returns, not a hardcoded literal. Slides (and
	// their rIds, threaded into SldIdLst) are added later by AddSlide.
	presRels := pkg.Relationships(PathPresentation)
	masterRID, err := presRels.Add(RelTypeSlideMaster, "slideMasters/slideMaster1.xml", "Internal")
	if err != nil {
		panic(err)
	}

	// SldIdLst starts nil: with zero slides added, encoding/xml writes
	// nothing for a nil pointer field, which is schema-valid (minOccurs=0).
	// AddSlide allocates it lazily on the first call. Marshaling happens
	// lazily too (only at Save/opc.Write time), so mutating *pres after
	// AddPart below is exactly what makes appending slides later work.
	pres := &XMLPresentation{
		XmlnsA: drawingml.NamespaceMain,
		XmlnsR: drawingml.NamespaceRelationships,
		XmlnsP: NamespaceMain,
		SldMasterIdLst: &SldMasterIdLst{
			Entries: []*SldMasterId{{ID: firstSldMasterID, RID: masterRID}},
		},
		SldSz:   &SldSz{Cx: cfg.slideWidthEMU, Cy: cfg.slideHeightEMU, Type: cfg.slideSizeType},
		NotesSz: &NotesSz{Cx: notesWidthEMU, Cy: notesHeightEMU},
	}
	pkg.AddPart(PathPresentation, ContentTypePresentation, pres)

	pkg.AddPart(PathCoreProps, opc.ContentTypeCoreProperties, NewCoreProperties("", ""))
	pkg.AddPart(PathAppProps, opc.ContentTypeExtendedProperties, NewAppProperties())

	rootRels := pkg.Relationships("")
	if _, err := rootRels.Add(RelTypeOfficeDocument, PathPresentation, "Internal"); err != nil {
		panic(err)
	}
	if _, err := rootRels.Add(opc.RelTypeCoreProperties, PathCoreProps, "Internal"); err != nil {
		panic(err)
	}
	if _, err := rootRels.Add(opc.RelTypeExtendedProperties, PathAppProps, "Internal"); err != nil {
		panic(err)
	}

	return &Presentation{pkg: pkg, pres: pres, presRels: presRels}
}

// AddSlide appends a new, empty slide — using the presentation's one slide
// layout — and returns a handle for adding shapes to it.
func (p *Presentation) AddSlide() *Slide {
	p.slideCount++
	n := p.slideCount
	path := SlidePath(n)

	spTree := NewEmptySpTree()
	cSld := &CSld{SpTree: spTree}
	p.pkg.AddPart(path, ContentTypeSlide, &XMLSlide{
		XmlnsA:    drawingml.NamespaceMain,
		XmlnsR:    drawingml.NamespaceRelationships,
		XmlnsP:    NamespaceMain,
		CSld:      cSld,
		ClrMapOvr: NewClrMapOvrInherit(),
	})
	if _, err := p.pkg.Relationships(path).Add(RelTypeSlideLayout, "../slideLayouts/slideLayout1.xml", "Internal"); err != nil {
		panic(err) // static, well-formed arguments; cannot fail
	}

	slideRID, err := p.presRels.Add(RelTypeSlide, "slides/slide"+strconv.Itoa(n)+".xml", "Internal")
	if err != nil {
		panic(err)
	}

	if p.pres.SldIdLst == nil {
		p.pres.SldIdLst = &SldIdLst{}
	}
	p.pres.SldIdLst.Entries = append(p.pres.SldIdLst.Entries, &SldId{
		ID:  firstSldID + uint32(n-1),
		RID: slideRID,
	})

	return &Slide{pres: p, path: path, cSld: cSld, spTree: spTree, nextShapeID: firstShapeID}
}

// addErr records a user-input validation error raised deep in a fluent
// chain (see text_builder.go), where returning early would break chaining.
// Save returns the first one recorded. Nil errs are ignored so call sites
// don't need their own nil check.
func (p *Presentation) addErr(err error) {
	if err != nil {
		p.errs = append(p.errs, err)
	}
}

// mediaBasename returns the ppt/media/ basename data should be embedded
// under, reusing an already-embedded part when byte-identical content was
// embedded before rather than writing a duplicate part every time an
// AddImage* call places the same image again (across slides, or several
// times on one slide). Slide.addPicture calls this instead of unconditionally
// allocating a new media part.
func (p *Presentation) mediaBasename(data []byte, contentType, ext string) string {
	hash := sha256.Sum256(data)
	if basename, ok := p.mediaByHash[hash]; ok {
		return basename
	}

	name := p.pkg.IDs().NextID("image")
	basename := name + "." + ext
	p.pkg.AddMediaPart("ppt/media/"+basename, contentType, data)

	if p.mediaByHash == nil {
		p.mediaByHash = make(map[[sha256.Size]byte]string)
	}
	p.mediaByHash[hash] = basename

	return basename
}

// Save writes the presentation to w as a .pptx file. If any fluent builder
// call recorded a validation error, Save returns the first one instead of
// writing.
func (p *Presentation) Save(w io.Writer) error {
	if len(p.errs) > 0 {
		return p.errs[0]
	}
	return p.pkg.Write(w)
}
