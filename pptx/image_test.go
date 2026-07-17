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
	"image"
	"image/color"
	"image/gif"
	"image/jpeg"
	"image/png"
	"testing"

	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/opc"
)

// solidImage returns an in-memory w x h RGBA image filled with a solid
// color, for encoding into whichever test format needs bytes to decode.
func solidImage(w, h int) *image.RGBA {
	img := image.NewRGBA(image.Rect(0, 0, w, h))
	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			img.Set(x, y, color.RGBA{R: 0x1F, G: 0x49, B: 0x7D, A: 0xFF})
		}
	}
	return img
}

func TestImageMeta_PNG(t *testing.T) {
	var buf bytes.Buffer
	if err := png.Encode(&buf, solidImage(120, 80)); err != nil {
		t.Fatalf("png.Encode: %v", err)
	}

	w, h, ct, ext, err := imageMeta(buf.Bytes())
	if err != nil {
		t.Fatalf("imageMeta: %v", err)
	}
	if w != 120 || h != 80 {
		t.Errorf("got %dx%d, want 120x80", w, h)
	}
	if ct != opc.ContentTypePNG {
		t.Errorf("got content type %q, want %q", ct, opc.ContentTypePNG)
	}
	if ext != "png" {
		t.Errorf("got ext %q, want png", ext)
	}
}

func TestImageMeta_JPEG(t *testing.T) {
	var buf bytes.Buffer
	if err := jpeg.Encode(&buf, solidImage(64, 48), nil); err != nil {
		t.Fatalf("jpeg.Encode: %v", err)
	}

	w, h, ct, ext, err := imageMeta(buf.Bytes())
	if err != nil {
		t.Fatalf("imageMeta: %v", err)
	}
	if w != 64 || h != 48 {
		t.Errorf("got %dx%d, want 64x48", w, h)
	}
	if ct != opc.ContentTypeJPEG {
		t.Errorf("got content type %q, want %q", ct, opc.ContentTypeJPEG)
	}
	if ext != "jpeg" {
		t.Errorf("got ext %q, want jpeg", ext)
	}
}

func TestImageMeta_GIF(t *testing.T) {
	var buf bytes.Buffer
	if err := gif.Encode(&buf, solidImage(32, 32), nil); err != nil {
		t.Fatalf("gif.Encode: %v", err)
	}

	w, h, ct, ext, err := imageMeta(buf.Bytes())
	if err != nil {
		t.Fatalf("imageMeta: %v", err)
	}
	if w != 32 || h != 32 {
		t.Errorf("got %dx%d, want 32x32", w, h)
	}
	if ct != opc.ContentTypeGIF {
		t.Errorf("got content type %q, want %q", ct, opc.ContentTypeGIF)
	}
	if ext != "gif" {
		t.Errorf("got ext %q, want gif", ext)
	}
}

func TestImageMeta_UnrecognizedDataErrors(t *testing.T) {
	_, _, _, _, err := imageMeta([]byte("not an image"))
	if err == nil {
		t.Fatal("expected an error for unrecognized image data")
	}
}

func TestEmuPerPixel96DPI_MatchesInchConversion(t *testing.T) {
	// 96 px at 96 DPI must be exactly 1 inch.
	if got := 96 * emuPerPixel96DPI; got != drawingml.EMUsPerInch {
		t.Errorf("96px at 96DPI = %d EMU, want %d (1 inch)", got, drawingml.EMUsPerInch)
	}
}
