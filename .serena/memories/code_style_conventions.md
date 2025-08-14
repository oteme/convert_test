# Code Style and Conventions

## Language and Encoding
- **Language**: Python 3.x
- **Encoding**: UTF-8 (essential for Japanese text processing)
- **Comments**: Mix of Japanese and English

## Naming Conventions

### Variables and Functions
- **Snake_case** for functions and variables: `parse_cell_line()`, `current_table`
- **Descriptive names**: `header_depth`, `manual_header_depth`, `keep_dividers`
- **Japanese terms preserved**: `sanitize_name()` keeps Japanese characters

### Constants and Regex
- **UPPERCASE** for regex patterns: `CAPTION_RE`, `TABLE_START_RE`, `ROW_RE`
- **Meaningful pattern names**: `CELL_RE`, `SPAN_RE`, `HEADER_HINT_RE`

## Code Organization

### Function Structure
- Functions are well-modularized with single responsibilities
- Clear separation between parsing, processing, and output functions
- Helper functions for common operations (`strip_inline_tags`, `sanitize_name`)

### Type Hints
- Consistent use of type hints from `typing` module
- `List`, `Tuple`, `Dict`, `Any` imports used throughout
- Function signatures include return types where applicable

### Documentation
- Docstrings in Japanese for main functions
- Inline comments explaining complex logic
- CLI help text in Japanese

## Error Handling
- Basic error handling for file operations
- Validation for command-line arguments (e.g., header-depth minimum value)
- Graceful handling of missing or malformed data

## Import Organization
- Standard library imports first: `sys`, `re`, `json`, `argparse`
- Type imports: `from typing import List, Tuple, Dict, Any`  
- Third-party imports: `from openpyxl import Workbook`

## String Processing
- Heavy use of regular expressions for pattern matching
- String methods: `.strip()`, `.splitlines()`, `.match()`, `.group()`
- Unicode-aware processing for Japanese text

## Data Structures
- Tuples for cell data: `(value, rowspan, colspan, is_header)`
- Dictionaries for table structure with consistent keys
- Lists for rows and table collections

## CLI Design
- `argparse` for command-line interface
- Sensible defaults for optional parameters
- Comprehensive help text and parameter descriptions