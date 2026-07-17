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

	// Only "External" requires the TargetMode attribute; for internal
	// relationships Office expects it to be omitted entirely.
	mode := strings.TrimSpace(targetMode)
	if strings.EqualFold(mode, "internal") || mode == "" {
		mode = ""
	}

	rm.relationships[id] = &Relationship{ID: id, Type: relType, Target: target, TargetMode: mode}
	return id, nil
}

// AddImage adds an image relationship (Internal, RelTypeImage).
func (rm *RelationshipManager) AddImage(target string) (string, error) {
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

// All returns every relationship owned by this manager.
func (rm *RelationshipManager) All() []*Relationship {
	rm.mu.RLock()
	defer rm.mu.RUnlock()

	rels := make([]*Relationship, 0, len(rm.relationships))
	for _, rel := range rm.relationships {
		rels = append(rels, rel)
	}
	return rels
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

	mode := strings.TrimSpace(targetMode)
	if strings.EqualFold(mode, "internal") || mode == "" {
		mode = ""
	}

	rm.relationships[id] = &Relationship{ID: id, Type: relType, Target: target, TargetMode: mode}

	if n, err := strconv.Atoi(strings.TrimPrefix(strings.ToLower(id), "rid")); err == nil {
		rm.ensureCounterAtLeast(uint64(n))
	}
	return nil
}

// ensureCounterAtLeast advances relCounter to at least value, without ever
// decreasing it, so IDs generated after RegisterExisting never collide with
// a preserved one.
func (rm *RelationshipManager) ensureCounterAtLeast(value uint64) {
	for {
		current := rm.relCounter.Load()
		if current >= value {
			return
		}
		if rm.relCounter.CompareAndSwap(current, value) {
			return
		}
	}
}

// ToXML converts the relationships to their .rels XML representation.
func (rm *RelationshipManager) ToXML() *XMLRelationships {
	rm.mu.RLock()
	defer rm.mu.RUnlock()

	out := &XMLRelationships{
		Xmlns:         NamespacePackageRels,
		Relationships: make([]*XMLRelationship, 0, len(rm.relationships)),
	}
	for _, rel := range rm.relationships {
		out.Relationships = append(out.Relationships, &XMLRelationship{
			ID:         rel.ID,
			Type:       rel.Type,
			Target:     rel.Target,
			TargetMode: rel.TargetMode,
		})
	}
	return out
}
