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

// Package errors provides structured error types for pptxgo.
package errors

import (
	"fmt"
	"strings"
)

// Error codes for categorizing errors.
const (
	ErrCodeValidation = "VALIDATION_ERROR"
	ErrCodeNotFound   = "NOT_FOUND"
	ErrCodeInternal   = "INTERNAL_ERROR"
)

// PptxError represents a structured error in pptxgo.
type PptxError struct {
	Code    string
	Op      string
	Err     error
	Message string
	Context map[string]interface{}
}

// Error implements the error interface.
func (e *PptxError) Error() string {
	var parts []string

	if e.Op != "" {
		parts = append(parts, fmt.Sprintf("operation=%s", e.Op))
	}
	if e.Code != "" {
		parts = append(parts, fmt.Sprintf("code=%s", e.Code))
	}
	if e.Message != "" {
		parts = append(parts, e.Message)
	}
	if e.Err != nil {
		parts = append(parts, fmt.Sprintf("cause=%v", e.Err))
	}
	if len(e.Context) > 0 {
		var ctx []string
		for k, v := range e.Context {
			ctx = append(ctx, fmt.Sprintf("%s=%v", k, v))
		}
		parts = append(parts, fmt.Sprintf("context={%s}", strings.Join(ctx, ", ")))
	}

	return strings.Join(parts, " | ")
}

// Unwrap returns the underlying error.
func (e *PptxError) Unwrap() error {
	return e.Err
}

// Is reports whether target is a PptxError with the same Code. This matches
// by category only: every validation error compares equal to every other,
// so it cannot single out one specific condition. That is fine while nothing
// discriminates errors this way. TODO: if sentinel errors are introduced for
// callers to match on, tighten this (e.g. compare identity or an added Kind)
// so distinct conditions sharing a Code stop comparing equal.
func (e *PptxError) Is(target error) bool {
	t, ok := target.(*PptxError)
	if !ok {
		return false
	}
	return e.Code == t.Code
}

// InvalidArgument creates a validation error for invalid arguments.
func InvalidArgument(op, field string, value interface{}, message string) error {
	return &PptxError{
		Code:    ErrCodeValidation,
		Op:      op,
		Message: fmt.Sprintf("field=%s, value=%v: %s", field, value, message),
	}
}

// NotFound creates a "not found" error.
func NotFound(op, item string) error {
	return &PptxError{
		Code:    ErrCodeNotFound,
		Op:      op,
		Message: fmt.Sprintf("%s not found", item),
	}
}

// Wrap wraps an error with operation context.
func Wrap(err error, op string) error {
	if err == nil {
		return nil
	}
	return &PptxError{
		Code: ErrCodeInternal,
		Op:   op,
		Err:  err,
	}
}
