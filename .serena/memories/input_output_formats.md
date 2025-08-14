# Input/Output Formats

## Input Format Specification

### File Structure
The input is a text file with custom markup tags for table structure:

```
<"表タイトル">
<"表001">
<"行">
<"セル" rowspan=1 colspan=2 header=true>Header Cell
<"セル" rowspan=2 colspan=1>Data Cell
<"行">
<"セル">Another Cell
```

### Markup Tags
- **`<"キャプション">`** - Table caption/title
- **`<"表ID">`** - Table identifier and start marker
- **`<"行">`** - Row separator
- **`<"セル" attributes>`** - Cell definition with:
  - `rowspan` - Number of rows the cell spans
  - `colspan` - Number of columns the cell spans  
  - `header` - Boolean indicating if cell is a header

### Text Processing
- Inline tags are stripped: `<popup>`, `<M>`, etc.
- Japanese medical/pharmaceutical terminology preserved
- Unicode characters fully supported

## Output Formats

### JSON Structure
```json
{
  "tables": [
    {
      "id": "表001",
      "name": "Table Name",
      "header_depth": 2,
      "columns": [
        {
          "path": ["Level1", "Level2"],
          "key": "column_key"
        }
      ],
      "data": [
        {
          "column_key": "value",
          "分類": "classification"
        }
      ]
    }
  ]
}
```

### Excel Structure
- Multi-sheet workbook
- Each table on separate sheet
- Hierarchical headers preserved
- Cell merging applied
- Japanese text fully supported

## Command Line Options

### JSON Conversion Options
- `--header-depth N` - Manual header depth specification
- `--value-policy {first_nonempty,last_nonempty,concat}` - Value extraction policy
- `--concat-sep "separator"` - Concatenation separator for merged cells
- `--add-classification` - Add classification columns from vertical spans
- `--group-key "key"` - Name for classification column
- `--nested` - Output nested JSON structure
- `--keep-dividers` - Preserve empty divider rows
- `--no-merge` - Skip automatic table merging

### Excel Conversion Options
- `--header-depth N` - Manual header depth specification
- Basic table structure preservation

## Data Types Handled
- Japanese text (UTF-8)
- Hierarchical table headers
- Cell spans (rowspan/colspan)
- Empty cells and divider rows
- Classification/grouping information
- Inline formatting tags (stripped in output)