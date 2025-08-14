# Project Overview

## Purpose
This is a Python-based table parsing and conversion tool designed to extract structured tables from specially formatted Japanese text files and convert them to JSON or Excel formats. The tool specifically handles:

- Text files with custom markup tags for table structure (e.g., `<"表名">`, `<"行">`, `<"セル">`)
- Hierarchical table headers with multiple levels
- Cell spanning (rowspan/colspan) information
- Japanese medical/pharmaceutical text processing (based on the input sample)

## Main Components

### Primary Scripts
1. **parse_tables_to_json.py** - Main parser that converts tables to JSON format
2. **parse_tables_to_xlsx.py** - Converts tables to Excel format using openpyxl

### Core Functionality
- Parses custom markup tags to identify table structure
- Handles hierarchical headers with configurable depth
- Supports cell merging (horizontal and vertical spans)  
- Provides classification columns for grouped data
- Offers various output formats (flat JSON, nested JSON, Excel)

## Tech Stack
- **Language**: Python 3.x
- **Dependencies**: openpyxl (for Excel output)
- **File Formats**: Input (custom tagged text), Output (JSON, XLSX)

## Input Format
The tool processes text files with special markup tags:
- `<"表名">` - Table captions
- `<"行">` - Row separators  
- `<"セル">` - Cell definitions with span information
- Inline tags for formatting and metadata

## Output Formats
1. **JSON**: Structured data with hierarchical headers and cell values
2. **Excel**: Multi-sheet workbooks with formatted tables