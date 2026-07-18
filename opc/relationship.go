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

package opc

import (
	"fmt"
	"sort"
	"strconv"
	"strings"
	"sync"
	"sync/atomic"

	"github.com/mmonterroca/pptxgo/pkg/errors"
)

// Relationship is a single OPC relationship: a typed, directed edge from an
// owning part (or the package root) to a target part or external resource.
type Relationship struct {
	ID         string // e.g. "rId1"
	Type       string // relationship type URI
	Target     string // target path (relative to the owner) or external URI
	TargetMode string // "External", or "" for Internal
}

// RelationshipManager owns the relationships for a single part (or, when
// used for the package root, for _rels/.rels). It is thread-safe.
//
// OPC scopes relationship IDs per owning part: rId1 in a slide's own .rels
// is unrelated to rId1 in another slide's .rels or in the root .rels. Every
// owner therefore gets its own manager with its own independent ID
// sequence — not a counter shared across the package, which would make
// every owner's IDs monotonically increase across the whole package instead
// of restarting at rId1, and is exactly the class of scope confusion that
// caused docxgo's per-part relationship resolution bug.
type RelationshipManager struct {
	mu            sync.RWMutex
	relationships map[string]*Relationship
	relCounter    atomic.Uint64
}

// NewRelationshipManager creates a new, independently-numbered relationship manager.
func NewRelationshipManager() *RelationshipManager {
	return &RelationshipManager{
		relationships: make(map[string]*Relationship, DefaultRelCapacity),
	}
}

// Add adds a new relationship and returns its generated ID.
//
// It does not deduplicate: adding the same target twice yields two distinct
// rIds and two <Relationship> entries. That is intentional here — a caller
// may legitimately want two independent relationship entries to the same
// target (e.g. two conceptually distinct hyperlinks that happen to point
// at the same URL) — but see AddImage, which does dedup, since two
// pictures on one slide sharing a media part should share one rId too.
func (rm *RelationshipManager) Add(relType, target, targetMode string) (string, error) {
	if relType == "" {
		return "", errors.InvalidArgument("RelationshipManager.Add", "relType", relType, "relationship type cannot be empty")
	}
	if target == "" {
		return "", errors.InvalidArgument("RelationshipManager.Add", "target", target, "target cannot be empty")
	}

	rm.mu.Lock()
	defer rm.mu.Unlock()

	id := fmt.Sprintf("rId%d", rm.relCounter.Add(1))
	rm.relationships[id] = &Relationship{ID: id, Type: relType, Target: target, TargetMode: normalizeTargetMode(targetMode)}
	return id, nil
}

// normalizeTargetMode canonicalizes a relationship's target mode. Only
// "External" is emitted; internal relationships (the default) omit the
// attribute entirely, which is what Office expects. Shared by Add and
// RegisterExisting so both entry points serialize an internal relationship
// identically.
func normalizeTargetMode(mode string) string {
	mode = strings.TrimSpace(mode)
	if strings.EqualFold(mode, "internal") || mode == "" {
		return ""
	}
	if strings.EqualFold(mode, "external") {
		return "External"
	}
	return mode
}

// AddImage adds an image relationship (Internal, RelTypeImage), reusing an
// existing one if this owner already has an image relationship to the same
// target. Without this, deduping the underlying media part (see
// pptx.Presentation's content-hash cache) would still leave every
// AddImage* call on the same slide creating its own redundant
// relationship to that one shared part.
func (rm *RelationshipManager) AddImage(target string) (string, error) {
	if rel, err := rm.GetByTarget(target); err == nil && rel.Type == RelTypeImage {
		return rel.ID, nil
	}
	return rm.Add(RelTypeImage, target, "Internal")
}

// AddHyperlink adds a hyperlink relationship (External, RelTypeHyperlink).
func (rm *RelationshipManager) AddHyperlink(target string) (string, error) {
	return rm.Add(RelTypeHyperlink, target, "External")
}

// Get retrieves a relationship by ID.
func (rm *RelationshipManager) Get(id string) (*Relationship, error) {
	rm.mu.RLock()
	defer rm.mu.RUnlock()

	rel, ok := rm.relationships[id]
	if !ok {
		return nil, errors.NotFound("RelationshipManager.Get", "relationship")
	}
	return rel, nil
}

// GetByTarget retrieves the first relationship matching the given target.
func (rm *RelationshipManager) GetByTarget(target string) (*Relationship, error) {
	rm.mu.RLock()
	defer rm.mu.RUnlock()

	for _, rel := range rm.relationships {
		if rel.Target == target {
			return rel, nil
		}
	}
	return nil, errors.NotFound("RelationshipManager.GetByTarget", "relationship")
}

// All returns every relationship owned by this manager, ordered by ascending
// numeric rId (see sortedLocked).
func (rm *RelationshipManager) All() []*Relationship {
	rm.mu.RLock()
	defer rm.mu.RUnlock()
	return rm.sortedLocked()
}

// sortedLocked returns the relationships ordered by ascending numeric rId.
// The manager stores relationships in a map (keyed by ID for lookup), whose
// iteration order Go randomizes; sorting here is what makes every .rels part
// — and therefore the whole .pptx byte stream — reproducible run-to-run,
// which reproducible builds and golden-file tests depend on. Must be called
// with rm.mu held (read or write).
func (rm *RelationshipManager) sortedLocked() []*Relationship {
	rels := make([]*Relationship, 0, len(rm.relationships))
	for _, rel := range rm.relationships {
		rels = append(rels, rel)
	}
	sort.Slice(rels, func(i, j int) bool {
		return relIDLess(rels[i].ID, rels[j].ID)
	})
	return rels
}

// relIDLess orders relationship IDs by their numeric suffix ("rId2" < "rId10"),
// falling back to a plain string comparison for any ID that does not fit the
// "rId<n>" shape.
func relIDLess(a, b string) bool {
	na, erra := strconv.ParseUint(strings.TrimPrefix(strings.ToLower(a), "rid"), 10, 64)
	nb, errb := strconv.ParseUint(strings.TrimPrefix(strings.ToLower(b), "rid"), 10, 64)
	if erra == nil && errb == nil && na != nb {
		return na < nb
	}
	// Falls through here whenever the numeric parts tie but the strings
	// don't (e.g. "rId1" vs "rId01", or two non-numeric IDs) — a comparator
	// that reports neither Less(a,b) nor Less(b,a) makes sort.Slice, which
	// is NOT a stable sort, free to order those elements differently across
	// calls. Comparing the raw strings breaks the tie deterministically.
	return a < b
}

// Count returns the number of relationships.
func (rm *RelationshipManager) Count() int {
	rm.mu.RLock()
	defer rm.mu.RUnlock()
	return len(rm.relationships)
}

// RegisterExisting adds a relationship with a caller-supplied ID (e.g. one
// loaded from an existing package's .rels part) without generating a new
// one, and advances the ID generator so future IDs never collide with it.
func (rm *RelationshipManager) RegisterExisting(id, relType, target, targetMode string) error {
	if id == "" {
		return errors.InvalidArgument("RelationshipManager.RegisterExisting", "id", id, "relationship id cannot be empty")
	}
	if relType == "" {
		return errors.InvalidArgument("RelationshipManager.RegisterExisting", "relType", relType, "relationship type cannot be empty")
	}
	if target == "" {
		return errors.InvalidArgument("RelationshipManager.RegisterExisting", "target", target, "relationship target cannot be empty")
	}

	rm.mu.Lock()
	defer rm.mu.Unlock()

	if _, exists := rm.relationships[id]; exists {
		return nil
	}

	rm.relationships[id] = &Relationship{ID: id, Type: relType, Target: target, TargetMode: normalizeTargetMode(targetMode)}

	if n, err := strconv.Atoi(strings.TrimPrefix(strings.ToLower(id), "rid")); err == nil {
		advanceAtLeast(&rm.relCounter, uint64(n))
	}
	return nil
}

// ToXML converts the relationships to their .rels XML representation.
func (rm *RelationshipManager) ToXML() *XMLRelationships {
	rm.mu.RLock()
	defer rm.mu.RUnlock()

	out := &XMLRelationships{
		Xmlns:         NamespacePackageRels,
		Relationships: make([]*XMLRelationship, 0, len(rm.relationships)),
	}
	for _, rel := range rm.sortedLocked() {
		out.Relationships = append(out.Relationships, &XMLRelationship{
			ID:         rel.ID,
			Type:       rel.Type,
			Target:     rel.Target,
			TargetMode: rel.TargetMode,
		})
	}
	return out
}
