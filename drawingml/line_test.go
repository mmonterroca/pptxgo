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

package drawingml

import (
	"strings"
	"testing"
)

func TestLn_MarshalsWidthAndSolidFill(t *testing.T) {
	ln := NewLn(Color{R: 0x1F, G: 0x49, B: 0x7D}, 12700)
	got := marshal(t, ln)

	if !strings.Contains(got, `<a:ln w="12700">`) {
		t.Errorf("expected a:ln with w attr, got %s", got)
	}
	if !strings.Contains(got, `<a:solidFill><a:srgbClr val="1F497D">`) {
		t.Errorf("expected nested solidFill, got %s", got)
	}
}

func TestLn_ZeroWidthOmitsAttr(t *testing.T) {
	ln := &Ln{}
	got := marshal(t, ln)
	if strings.Contains(got, `w=`) {
		t.Errorf("expected no w attr for zero width, got %s", got)
	}
}

func TestLn_PrstDashMarshalsAfterSolidFill(t *testing.T) {
	ln := NewLn(Color{R: 0x1F, G: 0x49, B: 0x7D}, 12700)
	ln.PrstDash = &PrstDash{Val: "dash"}
	got := marshal(t, ln)

	if !strings.Contains(got, `<a:prstDash val="dash">`) {
		t.Errorf("expected a:prstDash element, got %s", got)
	}
	fillIdx := strings.Index(got, "<a:solidFill>")
	dashIdx := strings.Index(got, "<a:prstDash")
	if fillIdx == -1 || dashIdx == -1 || fillIdx > dashIdx {
		t.Errorf("expected a:solidFill before a:prstDash (CT_LineProperties sequence), got %s", got)
	}
}
