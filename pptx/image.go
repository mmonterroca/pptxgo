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
	_ "image/gif"  // registers the GIF decoder used by image.DecodeConfig
	_ "image/jpeg" // registers the JPEG decoder used by image.DecodeConfig
	_ "image/png"  // registers the PNG decoder used by image.DecodeConfig

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
