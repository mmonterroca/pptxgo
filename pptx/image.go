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
	"image/jpeg"

	_ "image/gif" // registers the GIF decoder used by image.DecodeConfig
	_ "image/png" // registers the PNG decoder used by image.DecodeConfig

	"github.com/mmonterroca/pptxgo/drawingml"
	"github.com/mmonterroca/pptxgo/opc"
	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// emuPerPixel96DPI converts a pixel dimension to EMUs assuming the image
// renders at 96 DPI — PowerPoint's own assumption for an image with no
// explicit size. image.DecodeConfig reports pixel dimensions only, not a
// file's embedded density (e.g. a PNG's pHYs chunk or a JFIF header), so a
// higher-DPI asset auto-sizes larger than its physical print size; a caller
// who needs exact control should use AddImageWithSize instead.
const emuPerPixel96DPI = drawingml.EMUsPerInch / 96

// imageMeta decodes just the header of data — not the full image — to
// determine its pixel dimensions and format, and maps that format to the
// OPC content type and file extension pptxgo embeds it under. Supported
// formats are exactly those with a decoder in the Go standard library
// (PNG, JPEG, GIF); anything else — including BMP, TIFF, WMF, and EMF,
// which all have an opc.ContentType constant but no stdlib decoder —
// returns an error rather than silently mis-sizing or mis-typing the image.
//
// The reported dimensions are the stored pixel grid, exactly as decoded;
// EXIF orientation is applied by prepareImage (which physically rotates the
// pixels), not here, so imageMeta's width/height always match the bytes it
// was handed.
func imageMeta(data []byte) (wPx, hPx int, contentType, ext string, err error) {
	cfg, format, decodeErr := image.DecodeConfig(bytes.NewReader(data))
	if decodeErr != nil {
		return 0, 0, "", "", errors.InvalidArgument("imageMeta", "data", len(data),
			"not a recognized PNG, JPEG, or GIF image: "+decodeErr.Error())
	}

	switch format {
	case "png":
		return cfg.Width, cfg.Height, opc.ContentTypePNG, "png", nil
	case "jpeg":
		return cfg.Width, cfg.Height, opc.ContentTypeJPEG, "jpeg", nil
	case "gif":
		return cfg.Width, cfg.Height, opc.ContentTypeGIF, "gif", nil
	default:
		return 0, 0, "", "", errors.InvalidArgument("imageMeta", "format", format,
			"unsupported image format (only png, jpeg, gif are auto-detected)")
	}
}

// jpegReencodeQuality is the quality prepareImage re-encodes an
// orientation-corrected JPEG at. High enough that the extra generation of
// lossy compression a rotation forces is visually negligible, while keeping
// the file small.
const jpegReencodeQuality = 92

// prepareImage returns the bytes pptxgo should actually embed for data,
// along with that embedded image's pixel dimensions, content type, and
// extension. For every format except an EXIF-rotated JPEG it returns data
// unchanged (no decode/re-encode cost on the common path). For a JPEG whose
// EXIF Orientation tag is anything but "upright", it physically rotates and
// flips the pixels to match, then re-encodes: OOXML consumers (PowerPoint
// desktop, LibreOffice) render an embedded blip from its stored pixel grid
// and ignore EXIF orientation entirely, so baking the orientation into the
// pixels is the only way the image displays right — merely swapping the
// box's width and height would leave the sideways pixels stretched into the
// wrong-aspect box. The re-encoded JPEG carries no EXIF, so its orientation
// is unambiguously upright.
func prepareImage(data []byte) (out []byte, wPx, hPx int, contentType, ext string, err error) {
	wPx, hPx, contentType, ext, err = imageMeta(data)
	if err != nil {
		return nil, 0, 0, "", "", err
	}
	if ext != "jpeg" {
		return data, wPx, hPx, contentType, ext, nil
	}

	o := jpegOrientation(data)
	if o <= 1 || o > 8 {
		return data, wPx, hPx, contentType, ext, nil // already upright, or no readable tag
	}

	// The header parsed and declared a rotation; decode fully, apply it,
	// and re-encode. If the full decode or the re-encode fails, fall back
	// to embedding the original bytes (still a valid image, just not
	// orientation-corrected) rather than failing the whole AddImage call.
	src, decErr := jpeg.Decode(bytes.NewReader(data))
	if decErr != nil {
		return data, wPx, hPx, contentType, ext, nil
	}
	rotated := applyExifOrientation(src, o)
	var buf bytes.Buffer
	if encErr := jpeg.Encode(&buf, rotated, &jpeg.Options{Quality: jpegReencodeQuality}); encErr != nil {
		return data, wPx, hPx, contentType, ext, nil
	}
	b := rotated.Bounds()
	return buf.Bytes(), b.Dx(), b.Dy(), contentType, ext, nil
}

// applyExifOrientation returns src transformed so that an upright,
// EXIF-unaware viewer displays it the way the given EXIF orientation (1-8)
// intends. Orientations 5-8 include a 90/270-degree rotation and so
// transpose the dimensions; 2-4 only flip or rotate 180. Orientation 1 (or
// any out-of-range value) is returned unchanged.
func applyExifOrientation(src image.Image, orientation int) image.Image {
	if orientation <= 1 || orientation > 8 {
		return src
	}

	b := src.Bounds()
	w, h := b.Dx(), b.Dy()

	dstW, dstH := w, h
	if orientation >= 5 { // 5-8 rotate 90/270, transposing the axes
		dstW, dstH = h, w
	}
	dst := image.NewRGBA(image.Rect(0, 0, dstW, dstH))

	for y := 0; y < h; y++ {
		for x := 0; x < w; x++ {
			c := src.At(b.Min.X+x, b.Min.Y+y)
			var dx, dy int
			switch orientation {
			case 2: // mirror horizontal
				dx, dy = w-1-x, y
			case 3: // rotate 180
				dx, dy = w-1-x, h-1-y
			case 4: // mirror vertical
				dx, dy = x, h-1-y
			case 5: // transpose (mirror across the main diagonal)
				dx, dy = y, x
			case 6: // rotate 90 clockwise
				dx, dy = h-1-y, x
			case 7: // transverse (mirror across the anti-diagonal)
				dx, dy = h-1-y, w-1-x
			case 8: // rotate 90 counter-clockwise (270 clockwise)
				dx, dy = y, w-1-x
			}
			dst.Set(dx, dy, c)
		}
	}
	return dst
}
