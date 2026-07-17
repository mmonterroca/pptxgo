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

import "encoding/binary"

// exifOrientationDefault is EXIF's "normal, no rotation" orientation (1) —
// what a viewer with no EXIF support assumes, and what jpegOrientation
// returns whenever it can't find or parse an orientation tag.
const exifOrientationDefault = 1

// jpegOrientation returns the EXIF orientation (1-8) of JPEG data, or
// exifOrientationDefault if data carries no EXIF orientation tag, isn't a
// JPEG, or the EXIF block is malformed. It never errors — a photo with no
// EXIF metadata (or a synthetic, non-camera JPEG, like the ones this
// package's own tests generate) is common and entirely valid.
//
// image.DecodeConfig, which imageMeta otherwise relies on for pixel
// dimensions, reports the physical pixel grid only — it has no notion of
// EXIF's Orientation tag, so a photo shot in portrait but stored with its
// sensor's native (often landscape) pixel grid decodes with width and
// height swapped relative to how it displays. Orientations 5-8 (partly a
// 90/270-degree rotation) need that swap applied to the dimensions pptxgo
// hands to PowerPoint; 1-4 (identity or a simple flip) don't.
func jpegOrientation(data []byte) int {
	if len(data) < 4 || data[0] != 0xFF || data[1] != 0xD8 {
		return exifOrientationDefault
	}

	i := 2
	for i+4 <= len(data) {
		if data[i] != 0xFF {
			return exifOrientationDefault
		}
		marker := data[i+1]

		// Markers with no payload: standalone RST0-RST7 and SOI/EOI.
		if marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7) {
			i += 2
			continue
		}
		// Start of scan: image data follows: no more metadata markers.
		if marker == 0xDA {
			return exifOrientationDefault
		}

		segLen := int(data[i+2])<<8 | int(data[i+3])
		if segLen < 2 || i+2+segLen > len(data) {
			return exifOrientationDefault
		}
		payload := data[i+4 : i+2+segLen]

		if marker == 0xE1 && len(payload) > 6 && string(payload[:6]) == "Exif\x00\x00" {
			if o := parseExifOrientation(payload[6:]); o != 0 {
				return o
			}
			return exifOrientationDefault
		}
		i += 2 + segLen
	}
	return exifOrientationDefault
}

// parseExifOrientation reads the Orientation tag (0x0112) from a
// TIFF-structured EXIF block — the bytes immediately following the
// "Exif\0\0" marker in a JPEG APP1 segment. Returns 0 if the tag is absent
// or the block is too short/malformed to parse, letting the caller fall
// back to the default orientation rather than guessing.
func parseExifOrientation(tiff []byte) int {
	if len(tiff) < 8 {
		return 0
	}

	var bo binary.ByteOrder
	switch string(tiff[:2]) {
	case "II":
		bo = binary.LittleEndian
	case "MM":
		bo = binary.BigEndian
	default:
		return 0
	}
	if bo.Uint16(tiff[2:4]) != 0x002A {
		return 0
	}

	ifdOffset := int(bo.Uint32(tiff[4:8]))
	if ifdOffset < 0 || ifdOffset+2 > len(tiff) {
		return 0
	}
	entryCount := int(bo.Uint16(tiff[ifdOffset : ifdOffset+2]))
	base := ifdOffset + 2

	const orientationTag = 0x0112
	const entrySize = 12
	for e := 0; e < entryCount; e++ {
		off := base + e*entrySize
		if off+entrySize > len(tiff) {
			return 0
		}
		if bo.Uint16(tiff[off:off+2]) != orientationTag {
			continue
		}
		return int(bo.Uint16(tiff[off+8 : off+10]))
	}
	return 0
}
