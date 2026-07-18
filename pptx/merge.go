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
	"regexp"
	"sort"
	"strings"

	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// MergeData maps a placeholder name (the text between delimiters, e.g.
// "client_name" for the default "{{client_name}}") to its replacement text.
type MergeData = map[string]string

// MergeOption configures Template.Merge/OpenSlide.Merge.
type MergeOption func(*mergeConfig)

type mergeConfig struct {
	open, close string
	strict      bool
}

func newMergeConfig(opts []MergeOption) *mergeConfig {
	cfg := &mergeConfig{open: "{{", close: "}}"}
	for _, opt := range opts {
		opt(cfg)
	}
	return cfg
}

// WithDelimiters overrides Merge's placeholder delimiters (default "{{"/"}}").
func WithDelimiters(open, close string) MergeOption {
	return func(c *mergeConfig) { c.open, c.close = open, close }
}

// WithStrictMode makes Merge return an error if any placeholder found in
// the slide text has no matching key in the supplied data, instead of
// silently leaving that one placeholder untouched.
func WithStrictMode() MergeOption {
	return func(c *mergeConfig) { c.strict = true }
}

// placeholderPattern compiles a regexp matching "<open>key<close>" for the
// given delimiters, capturing key. The key is trimmed of surrounding
// whitespace — "{{ name }}" and "{{name}}" both match "name", a common
// real-world authoring habit worth tolerating rather than rejecting.
func placeholderPattern(open, close string) *regexp.Regexp {
	return regexp.MustCompile(regexp.QuoteMeta(open) + `\s*([^{}]+?)\s*` + regexp.QuoteMeta(close))
}

// Replace performs a literal (non-{{}}) substring replacement across every
// slide's text, returning the number of occurrences replaced. See
// substitute.go for how a match split across runs by PowerPoint's own
// editing history is healed before matching.
func (t *Template) Replace(old, new string) (int, error) {
	total := 0
	for _, pth := range t.slidePaths {
		n, err := t.replaceInSlide(pth, old, new)
		if err != nil {
			return total, err
		}
		total += n
	}
	return total, nil
}

// Replace performs a literal substring replacement on just this slide.
func (s *OpenSlide) Replace(old, new string) (int, error) {
	return s.tmpl.replaceInSlide(s.path, old, new)
}

func (t *Template) replaceInSlide(pth, old, new string) (int, error) {
	raw, err := t.rawSlideBytes(pth)
	if err != nil {
		return 0, err
	}

	count := 0
	out, changed, err := substituteSlideText(raw, func(text string) string {
		n := strings.Count(text, old)
		if n == 0 {
			return text
		}
		count += n
		return strings.ReplaceAll(text, old, new)
	})
	if err != nil {
		return 0, err
	}
	if changed > 0 {
		if err := t.writeSlideBytes(pth, out); err != nil {
			return count, err
		}
	}
	return count, nil
}

// Merge substitutes every placeholder (default "{{key}}", see
// WithDelimiters) found across all slides with data[key], returning the
// number of placeholders substituted. With WithStrictMode, a placeholder
// with no matching key in data makes Merge return an error (the
// substitutions that DID match are still applied) instead of silently
// leaving unmatched placeholders in place.
func (t *Template) Merge(data MergeData, opts ...MergeOption) (int, error) {
	cfg := newMergeConfig(opts)
	pattern := placeholderPattern(cfg.open, cfg.close)

	total := 0
	var missing []string
	for _, pth := range t.slidePaths {
		raw, err := t.rawSlideBytes(pth)
		if err != nil {
			return total, err
		}
		n, out, changed, err := mergeSlide(raw, pattern, data, &missing)
		if err != nil {
			return total, err
		}
		total += n
		if changed > 0 {
			if err := t.writeSlideBytes(pth, out); err != nil {
				return total, err
			}
		}
	}
	if cfg.strict && len(missing) > 0 {
		return total, missingPlaceholdersErr(missing)
	}
	return total, nil
}

// Merge substitutes placeholders on just this slide — see Template.Merge.
func (s *OpenSlide) Merge(data MergeData, opts ...MergeOption) (int, error) {
	cfg := newMergeConfig(opts)
	pattern := placeholderPattern(cfg.open, cfg.close)

	raw, err := s.tmpl.rawSlideBytes(s.path)
	if err != nil {
		return 0, err
	}
	var missing []string
	n, out, changed, err := mergeSlide(raw, pattern, data, &missing)
	if err != nil {
		return 0, err
	}
	if changed > 0 {
		if err := s.tmpl.writeSlideBytes(s.path, out); err != nil {
			return n, err
		}
	}
	if cfg.strict && len(missing) > 0 {
		return n, missingPlaceholdersErr(missing)
	}
	return n, nil
}

func missingPlaceholdersErr(missing []string) error {
	sort.Strings(missing)
	return errors.InvalidArgument("Merge", "data", missing, "placeholder(s) with no matching key: "+strings.Join(missing, ", "))
}

// mergeSlide applies pattern-based substitution to one slide's raw bytes,
// appending any unmatched placeholder key to *missing (a shared
// accumulator across every slide in a Template.Merge call, so strict mode
// reports every miss in the whole presentation, not just the first).
func mergeSlide(raw []byte, pattern *regexp.Regexp, data MergeData, missing *[]string) (count int, out []byte, changed int, err error) {
	out, changed, err = substituteSlideText(raw, func(text string) string {
		return pattern.ReplaceAllStringFunc(text, func(match string) string {
			key := pattern.FindStringSubmatch(match)[1]
			val, ok := data[key]
			if !ok {
				*missing = append(*missing, key)
				return match
			}
			count++
			return val
		})
	})
	return count, out, changed, err
}

// PlaceholderNames returns the distinct placeholder names (default
// "{{key}}" delimiters) found anywhere across all slides, for inspecting
// what a template expects before calling Merge. Reuses extractText
// (open.go) rather than the run-grouping machinery above — finding names is
// pure reading, and concatenating a:t content in document order already
// heals a run-split pattern the same way Text() does, with no need to track
// per-run byte spans for a read-only pass.
func (t *Template) PlaceholderNames() ([]string, error) {
	pattern := placeholderPattern("{{", "}}")
	seen := make(map[string]bool)

	for _, pth := range t.slidePaths {
		raw, err := t.rawSlideBytes(pth)
		if err != nil {
			return nil, err
		}
		text, err := extractText(raw)
		if err != nil {
			return nil, err
		}
		for _, m := range pattern.FindAllStringSubmatch(text, -1) {
			seen[m[1]] = true
		}
	}

	names := make([]string, 0, len(seen))
	for k := range seen {
		names = append(names, k)
	}
	sort.Strings(names)
	return names, nil
}
