# Suggested Commands

## Development Commands

### Python Environment
```bash
# Activate virtual environment
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Deactivate virtual environment
deactivate
```

### Running the Scripts

#### JSON Conversion
```bash
# Basic usage
python parse_tables_to_json.py input.txt output.json

# With options
python parse_tables_to_json.py input.txt output.json --header-depth 2 --add-classification --nested

# Full feature example
python parse_tables_to_json.py input.txt output.json \
  --header-depth 3 \
  --value-policy concat \
  --concat-sep " | " \
  --add-classification \
  --group-key "分類" \
  --nested \
  --keep-dividers
```

#### Excel Conversion  
```bash
# Basic usage
python parse_tables_to_xlsx.py input.txt output.xlsx

# With header depth
python parse_tables_to_xlsx.py input.txt output.xlsx --header-depth 2
```

## System Commands

### File Operations
```bash
# List files
ls -la

# View file contents
cat input.txt | head -50

# Search in files
grep -n "表" input.txt

# Find files
find . -name "*.py" -type f
```

### Git Operations
```bash
# Initialize repository
git init

# Add files
git add .

# Commit changes  
git commit -m "Initial commit"

# Check status
git status

# View differences
git diff
```

## Testing Commands

### Manual Testing
```bash
# Test with sample input
python parse_tables_to_json.py input.txt test_output.json

# Validate JSON output
python -m json.tool test_output.json

# Compare outputs
diff output1.json output2.json
```

### Debugging
```bash
# Run with Python debugger
python -m pdb parse_tables_to_json.py input.txt debug.json

# Check Python syntax
python -m py_compile parse_tables_to_json.py
```