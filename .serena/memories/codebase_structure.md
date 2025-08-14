# Codebase Structure

## File Organization

### Core Files
- `parse_tables_to_json.py` - Main JSON conversion script (756 lines)
- `parse_tables_to_xlsx.py` - Excel conversion script  
- `requirements.txt` - Single dependency: openpyxl
- `input.txt` - Sample input file with Japanese medical text and markup
- `.gitignore` - Standard Python gitignore with venv/, __pycache__/, etc.

### Data Files
Multiple output files present:
- `output*.json` - JSON output samples (1-5)
- `output*.xlsx` - Excel output samples (1-5)  
- `test_merge*.json` - Test files for table merging functionality

### Virtual Environment
- `venv/` - Python virtual environment with openpyxl installed

## Key Functions in parse_tables_to_json.py

### Parsing Functions
- `parse_text_to_tables()` - Main text parsing function
- `parse_cell_line()` - Individual cell parsing
- `parse_span()` - Cell span information parsing
- `strip_inline_tags()` - Tag cleaning utility

### Table Processing
- `merge_table_groups()` - Combine related tables
- `merge_tables_horizontally()` - Horizontal table merging
- `build_logical_columns()` - Create hierarchical column structure
- `fill_merged_labels()` - Handle merged cell labels

### Output Generation
- `table_to_json()` - Main JSON conversion
- `main()` - CLI interface with argparse

## Regular Expressions
The parser uses several regex patterns:
- `CAPTION_RE` - Table captions
- `TABLE_START_RE` - Table start markers  
- `ROW_RE` - Row separators
- `CELL_RE` - Cell definitions
- `SPAN_RE` - Cell span parsing
- `HEADER_HINT_RE` - Header detection