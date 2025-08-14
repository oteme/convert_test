# Task Completion Checklist

## When Task is Completed

### Code Quality Checks
Since this project doesn't have explicit linting/formatting tools configured, perform these manual checks:

```bash
# Check Python syntax
python -m py_compile parse_tables_to_json.py
python -m py_compile parse_tables_to_xlsx.py

# Test basic functionality
python parse_tables_to_json.py input.txt test_output.json
python parse_tables_to_xlsx.py input.txt test_output.xlsx
```

### Testing Requirements
1. **Functionality Testing**
   - Test with sample input file
   - Verify JSON output is valid: `python -m json.tool output.json`
   - Check Excel file opens correctly
   - Validate Japanese character encoding

2. **Edge Case Testing**
   - Empty input files
   - Files with no tables
   - Tables with complex cell spans
   - Various header depth configurations

3. **Command Line Interface Testing**
   - Test all CLI options
   - Verify help text: `python parse_tables_to_json.py --help`
   - Test invalid arguments handling

### File Management
1. **Clean up temporary files**
   ```bash
   rm -f test_*.json test_*.xlsx debug_*.txt
   ```

2. **Verify output files**
   - Check file sizes are reasonable
   - Ensure no corruption in Excel files
   - Validate JSON structure

### Documentation
1. **Update comments** if code logic changed
2. **Verify docstrings** are accurate
3. **Update CLI help text** if new options added

### Version Control (if applicable)
```bash
git add .
git status
git commit -m "Descriptive commit message"
```

## Quality Gates
- [ ] Code compiles without syntax errors
- [ ] Basic functionality works with sample input
- [ ] JSON output is valid
- [ ] Excel output opens correctly
- [ ] Japanese characters display properly
- [ ] CLI help is accurate
- [ ] No temporary/debug files left behind

## Known Limitations
- No automated testing framework
- No code formatting tools (black, autopep8)
- No static analysis tools (pylint, mypy)
- Limited error handling for malformed input