## Purpose

This capability lets users convert a speaking-order Excel file and a batch of PDF/DOCX materials into a reviewed, ordered, safely renamed ZIP package with a matching report.

## ADDED Requirements

### Requirement: Order Excel parsing
The system SHALL parse a `.xlsx` speaking-order file from a selected worksheet and header row into order items containing order number, title, code, and original Excel row number.

#### Scenario: Automatic column mapping succeeds
- **WHEN** the worksheet contains supported aliases for order number, title, and code
- **THEN** the system returns parsed order items using the detected columns

#### Scenario: Manual column mapping is required
- **WHEN** required columns cannot be detected automatically
- **THEN** the system allows the user or API caller to provide explicit column mappings before confirmation

#### Scenario: Duplicate order numbers are rejected
- **WHEN** parsed order items contain duplicate order numbers
- **THEN** the system blocks confirmation and reports the conflicting source rows

#### Scenario: Empty identity row is invalid
- **WHEN** a row has an order number but both title and code are empty
- **THEN** the system marks that row invalid and reports its source row

### Requirement: Document identity extraction
The system SHALL derive each uploaded PDF/DOCX document's title, material code, speaker or department, source snippet, and extraction error if any.

#### Scenario: Filename identity is available
- **WHEN** a filename contains a recognizable title or material code
- **THEN** the system uses that information as an initial document identity candidate

#### Scenario: Text extraction succeeds
- **WHEN** a document has extractable text or successful OCR output
- **THEN** the system uses the extracted text to populate or improve document identity fields

#### Scenario: Model extraction succeeds
- **WHEN** Ollama returns valid JSON for document identity extraction
- **THEN** the system records title, code, speaker or department, and source snippet without guessing missing values

#### Scenario: Single document extraction fails
- **WHEN** one document cannot be read, OCRed, or parsed by the model
- **THEN** the system records the failure for that document and continues processing the remaining documents

### Requirement: Deterministic matching and scoring
The system SHALL score every order item against candidate documents using code-first and title-second matching rules.

#### Scenario: Code exact match
- **WHEN** an order item code exactly equals a document code
- **THEN** the match score is 1.00 and the match method identifies exact code matching

#### Scenario: Normalized code match
- **WHEN** codes differ only by case, full-width characters, spaces, hyphens, or common punctuation
- **THEN** the match score is 0.98 and the match method identifies normalized code matching

#### Scenario: Title exact match
- **WHEN** an order title exactly equals a document title
- **THEN** the match score is 0.95 and the match method identifies exact title matching

#### Scenario: Title containment match
- **WHEN** a normalized order title contains or is contained by a normalized document title
- **THEN** the match score is 0.90 and the match method identifies title containment matching

#### Scenario: Title similarity match
- **WHEN** normalized title similarity is at least 0.80
- **THEN** the system assigns a title similarity score according to the configured similarity bands

### Requirement: One-to-one assignment
The system SHALL enforce that each order item has at most one selected document and each document is selected by at most one order item.

#### Scenario: Automatic assignment is unique
- **WHEN** the highest candidate score is at least 0.85, exceeds the second-best score by at least 0.10, and the candidate is unused
- **THEN** the system automatically assigns that document to the order item

#### Scenario: Ambiguous assignment requires review
- **WHEN** multiple candidates are close, low-confidence, or already occupied
- **THEN** the system marks the order item as needing manual review

#### Scenario: Duplicate manual assignment is detected
- **WHEN** a user manually selects a document already assigned to another order item
- **THEN** the system reports the conflict and blocks final export until resolved

### Requirement: Manual review and confirmation
The system SHALL allow users to review ambiguous, unmatched, failed, or manually changed matches before export.

#### Scenario: Review details are visible
- **WHEN** an order match needs review
- **THEN** the UI and API expose order details, candidate filenames, extracted identity fields, score, reason, source snippet, and occupancy status

#### Scenario: User confirms a candidate
- **WHEN** a user manually selects and confirms a document for an order item
- **THEN** the match status becomes confirmed and the selection participates in conflict checks and export

#### Scenario: User skips an item
- **WHEN** a user marks an order item as skipped
- **THEN** the system excludes that item from renamed document output while keeping it in the matching report

### Requirement: ZIP package export
The system SHALL export a ZIP containing safely renamed matched documents and `匹配清单.xlsx` without modifying original uploaded files.

#### Scenario: Ordered file names are generated
- **WHEN** a confirmed or automatically matched item is exported
- **THEN** the ZIP contains the source file bytes under `{order_no:03d}_{title}{original_extension}` after safe filename sanitization

#### Scenario: Duplicate output filenames are handled
- **WHEN** two output filenames would be identical after sanitization
- **THEN** the system makes them unique without changing extensions or allowing path traversal

#### Scenario: Matching report is included
- **WHEN** a ZIP is exported
- **THEN** the ZIP contains `匹配清单.xlsx` with order, source file, extracted identity, output filename, method, score, status, source snippet, and manual-change fields

#### Scenario: Original files remain unchanged
- **WHEN** the ZIP export completes
- **THEN** the uploaded source file names and source bytes remain unchanged in task storage

### Requirement: Ollama usage and failure behavior
The system SHALL use Ollama only for document identity extraction and limited ambiguous candidate resolution, with the server URL read from `OLLAMA_BASE_URL`.

#### Scenario: Ollama is unavailable
- **WHEN** Ollama cannot be reached or times out
- **THEN** the affected document or semantic decision records a clear error and deterministic matching continues where possible

#### Scenario: Ollama returns invalid JSON
- **WHEN** the model response cannot be parsed as valid JSON
- **THEN** the system records the model parse error and does not treat guessed content as confirmed identity data

### Requirement: Temporary file lifecycle
The system SHALL keep uploaded task files available for retry and export until the task is explicitly cleaned up.

#### Scenario: Failed file can be retried
- **WHEN** a document extraction fails
- **THEN** the original uploaded bytes or temporary file remain available for a retry

#### Scenario: Cleanup removes task files
- **WHEN** a task is explicitly cleaned up
- **THEN** temporary files owned by that task are removed without affecting unrelated files
