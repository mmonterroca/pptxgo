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
	"encoding/xml"
	"testing"

	"github.com/mmonterroca/pptxgo/opc"
)

func TestOfficeDocumentRelsNS_MatchesOpcConstant(t *testing.T) {
	// officeDocumentRelsNS is a separate literal (not a reference to
	// opc.NamespaceOfficeDocumentRels) purely because nav.go's original
	// design tried expressing this as a struct tag, which requires a
	// compile-time literal -- see xmlSldIdNav's doc comment for why that
	// approach was abandoned. This pins the two together so a future change
	// to one doesn't silently desync from the other.
	if opc.NamespaceOfficeDocumentRels != officeDocumentRelsNS {
		t.Fatalf("opc.NamespaceOfficeDocumentRels = %q, want %q (update nav.go's officeDocumentRelsNS to match)",
			opc.NamespaceOfficeDocumentRels, officeDocumentRelsNS)
	}
}

func TestXmlSldIdNav_UnmarshalsIDAndRIDDistinctly(t *testing.T) {
	// The whole point of the namespace-qualified tag: "id" (no namespace)
	// and "r:id" (officeDocument relationships namespace) share a local
	// name and must not collide.
	docXML := `<presentation xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
		<sldIdLst>
			<sldId id="256" r:id="rId7"/>
			<sldId id="257" r:id="rId8"/>
		</sldIdLst>
	</presentation>`

	var nav xmlPresentationNav
	if err := xml.Unmarshal([]byte(docXML), &nav); err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if nav.SldIdLst == nil {
		t.Fatal("expected SldIdLst to be populated")
	}
	if len(nav.SldIdLst.Entries) != 2 {
		t.Fatalf("expected 2 sldId entries, got %d", len(nav.SldIdLst.Entries))
	}
	if nav.SldIdLst.Entries[0].ID != 256 || nav.SldIdLst.Entries[0].RID != "rId7" {
		t.Errorf("entry 0 = %+v, want ID=256 RID=rId7", nav.SldIdLst.Entries[0])
	}
	if nav.SldIdLst.Entries[1].ID != 257 || nav.SldIdLst.Entries[1].RID != "rId8" {
		t.Errorf("entry 1 = %+v, want ID=257 RID=rId8", nav.SldIdLst.Entries[1])
	}
}
