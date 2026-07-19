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
//
// The captured key excludes every distinct character actually used in open
// and close (not a hardcoded "{}"), so custom delimiters get the same
// protection the default braces have: without it, an unclosed placeholder
// followed by a real one using the same delimiter characters — e.g.
// "[[key1 ... [[key2]]" under WithDelimiters("[[", "]]") — would span both
// and capture "key1 ... [[key2" as one bogus key instead of stopping at the
// first close.
//
// Empty delimiters are rejected with an error rather than compiled: an
// empty open or close would build the character class "[^...]" with nothing
// (or only the other delimiter's characters) inside, and in the all-empty
// case the invalid regex "[^]" — which regexp.MustCompile would PANIC on,
// unwinding out of the public Merge/PlaceholderNames APIs. Returning an
// error (and using Compile, not MustCompile) keeps a misconfigured
// delimiter a normal error return instead of a crash.
func placeholderPattern(open, close string) (*regexp.Regexp, error) {
	if open == "" || close == "" {
		return nil, errors.InvalidArgument("WithDelimiters", "open/close", [2]string{open, close}, "placeholder delimiters must both be non-empty")
	}
	return regexp.Compile(regexp.QuoteMeta(open) + `\s*([^` + deduplicatedCharClass(open+close) + `]+?)\s*` + regexp.QuoteMeta(close))
}

// deduplicatedCharClass returns s's distinct runes, each escaped as needed
// for safe use inside a regexp character class ([...]): "]", "^", "-", and
// "\" are the only characters with special meaning inside a class, but all
// four are escaped unconditionally (regardless of position) rather than
// tracking the position-dependent RE2 rules for exactly when "^"/"-" are
// literal — over-escaping a character that didn't strictly need it is
// harmless, whereas under-escaping "]" would silently truncate the class.
func deduplicatedCharClass(s string) string {
	seen := make(map[rune]bool)
	var b strings.Builder
	for _, r := range s {
		if seen[r] {
			continue
		}
		seen[r] = true
		switch r {
		case '\\', ']', '^', '-':
			b.WriteRune('\\')
		}
		b.WriteRune(r)
	}
	return b.String()
}

// Replace performs a literal (non-{{}}) substring replacement across every
// slide's text, returning the number of occurrences replaced. An empty old
// is rejected with an error: strings.ReplaceAll treats "" as matching
// between every character, so an empty old (e.g. from an unset variable)
// would otherwise inject new throughout every run and silently corrupt the
// whole deck. See substitute.go for how a match split across runs by
// PowerPoint's own editing history is healed before matching.
func (t *Template) Replace(old, new string) (int, error) {
	if old == "" {
		return 0, errors.InvalidArgument("Replace", "old", old, "must not be empty")
	}
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

// Replace performs a literal substring replacement on just this slide — see
// Template.Replace, including why an empty old is an error.
func (s *OpenSlide) Replace(old, new string) (int, error) {
	if old == "" {
		return 0, errors.InvalidArgument("Replace", "old", old, "must not be empty")
	}
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
	pattern, err := placeholderPattern(cfg.open, cfg.close)
	if err != nil {
		return 0, err
	}

	total := 0
	var missing []string
	for _, pth := range t.slidePaths {
		n, err := t.mergeInSlide(pth, pattern, data, &missing)
		if err != nil {
			return total, err
		}
		total += n
	}
	if cfg.strict && len(missing) > 0 {
		return total, missingPlaceholdersErr(missing)
	}
	return total, nil
}

// Merge substitutes placeholders on just this slide — see Template.Merge.
func (s *OpenSlide) Merge(data MergeData, opts ...MergeOption) (int, error) {
	cfg := newMergeConfig(opts)
	pattern, err := placeholderPattern(cfg.open, cfg.close)
	if err != nil {
		return 0, err
	}

	var missing []string
	n, err := s.tmpl.mergeInSlide(s.path, pattern, data, &missing)
	if err != nil {
		return n, err
	}
	if cfg.strict && len(missing) > 0 {
		return n, missingPlaceholdersErr(missing)
	}
	return n, nil
}

// mergeInSlide runs pattern-based substitution on the single slide at pth
// and writes the result back if anything changed — the shared body of both
// Template.Merge (called once per slide) and OpenSlide.Merge (called once),
// mirroring how Replace routes both its levels through replaceInSlide, so
// the two Merge entry points can never silently diverge. The strict-mode
// check stays with the callers, since Template.Merge aggregates missing
// keys across every slide before deciding.
func (t *Template) mergeInSlide(pth string, pattern *regexp.Regexp, data MergeData, missing *[]string) (int, error) {
	raw, err := t.rawSlideBytes(pth)
	if err != nil {
		return 0, err
	}
	n, out, changed, err := mergeSlide(raw, pattern, data, missing)
	if err != nil {
		return 0, err
	}
	if changed > 0 {
		if err := t.writeSlideBytes(pth, out); err != nil {
			return n, err
		}
	}
	return n, nil
}

// missingPlaceholdersErr reports each distinct missing key once. missing
// naturally contains one entry per OCCURRENCE (a key repeated across
// several slides, or several times on one slide, appends once per
// occurrence — see mergeSlide), so it is deduplicated here rather than at
// every append site.
func missingPlaceholdersErr(missing []string) error {
	seen := make(map[string]bool, len(missing))
	deduped := make([]string, 0, len(missing))
	for _, m := range missing {
		if !seen[m] {
			seen[m] = true
			deduped = append(deduped, m)
		}
	}
	sort.Strings(deduped)
	return errors.InvalidArgument("Merge", "data", deduped, "placeholder(s) with no matching key: "+strings.Join(deduped, ", "))
}

// mergeSlide applies pattern-based substitution to one slide's raw bytes,
// appending any unmatched placeholder key to *missing (a shared
// accumulator across every slide in a Template.Merge call, so strict mode
// reports every miss in the whole presentation, not just the first).
//
// The per-group transform matches the pattern ONCE with
// FindAllStringSubmatchIndex and splices manually, rather than
// ReplaceAllStringFunc plus a second FindStringSubmatch inside the callback
// to recover the capture group — the callback form re-runs the engine on
// every match purely to extract a key the first match already found.
func mergeSlide(raw []byte, pattern *regexp.Regexp, data MergeData, missing *[]string) (count int, out []byte, changed int, err error) {
	out, changed, err = substituteSlideText(raw, func(text string) string {
		matches := pattern.FindAllStringSubmatchIndex(text, -1)
		if len(matches) == 0 {
			return text
		}
		var b strings.Builder
		last := 0
		for _, m := range matches {
			// m[0]:m[1] is the whole placeholder; m[2]:m[3] is the key
			// capture group.
			b.WriteString(text[last:m[0]])
			key := text[m[2]:m[3]]
			if val, ok := data[key]; ok {
				b.WriteString(val)
				count++
			} else {
				*missing = append(*missing, key)
				b.WriteString(text[m[0]:m[1]]) // leave the placeholder untouched
			}
			last = m[1]
		}
		b.WriteString(text[last:])
		return b.String()
	})
	return count, out, changed, err
}

// PlaceholderNames returns the distinct placeholder names found anywhere
// across all slides, for inspecting what a template expects before calling
// Merge. Delimiters default to "{{"/"}}", the same as Merge; pass
// WithDelimiters here too if Merge will be called with a custom pair, so
// PlaceholderNames reports what Merge will actually match (WithStrictMode
// is accepted but has no effect — there is nothing to be strict about when
// only listing names).
//
// Runs the SAME scan+run-grouping as Merge (via slideRunGroupTexts), NOT
// extractText: extractText concatenates every a:t in a paragraph regardless
// of formatting, so it would report a placeholder whose middle is
// separately formatted (e.g. "{{" and "}}" plain but the key bold — three
// runs with differing a:rPr that groupRuns keeps separate). Merge could
// never substitute such a placeholder (its regex runs per format group), so
// reporting it here would break the documented promise that
// PlaceholderNames lists exactly what Merge substitutes — and, worse, would
// make strict Merge look like it succeeded while the literal placeholder
// survived in the deck.
func (t *Template) PlaceholderNames(opts ...MergeOption) ([]string, error) {
	cfg := newMergeConfig(opts)
	pattern, err := placeholderPattern(cfg.open, cfg.close)
	if err != nil {
		return nil, err
	}
	seen := make(map[string]bool)

	for _, pth := range t.slidePaths {
		raw, err := t.rawSlideBytes(pth)
		if err != nil {
			return nil, err
		}
		texts, err := slideRunGroupTexts(raw)
		if err != nil {
			return nil, err
		}
		for _, text := range texts {
			for _, m := range pattern.FindAllStringSubmatch(text, -1) {
				seen[m[1]] = true
			}
		}
	}

	names := make([]string, 0, len(seen))
	for k := range seen {
		names = append(names, k)
	}
	sort.Strings(names)
	return names, nil
}

// slideRunGroupTexts returns the concatenated text of each run-group on the
// slide — the exact same units Merge's substitution operates on (a match
// must live within one group). Sharing this with PlaceholderNames is what
// keeps "what PlaceholderNames reports" and "what Merge substitutes"
// identical by construction.
func slideRunGroupTexts(raw []byte) ([]string, error) {
	runs, err := scanRuns(raw)
	if err != nil {
		return nil, err
	}
	groups := groupRuns(runs)
	texts := make([]string, len(groups))
	for i, g := range groups {
		texts[i] = g.text
	}
	return texts, nil
}
