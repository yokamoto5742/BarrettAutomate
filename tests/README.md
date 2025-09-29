# BarrettAutomate Test Suite

This directory contains comprehensive unit tests for the BarrettAutomate project, specifically testing the `BarrettCalculator` class and its functionality.

## Test Structure

### Files
- `test_barrett_calculator.py` - Main test file with comprehensive test cases
- `conftest.py` - Pytest configuration and shared fixtures
- `__init__.py` - Makes the tests directory a Python package

### Test Coverage

The test suite covers:

1. **Initialization Tests**
   - Valid file initialization
   - Custom headless setting
   - Path handling and URL configuration

2. **Data Loading Tests** (`load_patient_data`)
   - Successful data loading from Excel
   - File not found error handling
   - Pandas read errors

3. **Data Saving Tests** (`save_patient_data`)
   - Successful data saving
   - Backup file creation when results file exists
   - Write error handling

4. **Data Input Tests** (`input_patient_data`)
   - Successful patient data input to web form
   - Handling missing/empty fields
   - Exception handling during input

5. **Calculation Tests** (`calculate_and_get_result`)
   - Successful calculation and result extraction
   - Calculate button not found scenarios
   - Universal formula tab not found scenarios
   - Exception handling

6. **Result Extraction Tests**
   - Table-based extraction (`_extract_refraction_from_table`)
   - Alternative extraction methods (`_extract_refraction_alternative`)
   - Negative values, IOL power tolerance, malformed data

7. **Integration Tests** (`process_all_patients`)
   - Full workflow success scenarios
   - Input errors, calculation errors
   - Patient-level exceptions
   - Data loading errors

8. **Edge Cases**
   - Empty Excel files
   - NaN/null values in data
   - Negative refraction values
   - IOL power matching tolerance
   - Malformed table data

9. **Parametrized Tests**
   - IOL power matching logic
   - Refraction value extraction patterns

## Running Tests

### Basic Test Execution
```bash
# Activate virtual environment
source .venv/Scripts/activate  # On Windows: .venv\Scripts\activate

# Install test dependencies
pip install pytest pytest-mock pytest-cov

# Run all tests
python -m pytest tests/

# Run with verbose output
python -m pytest tests/ -v

# Run with short traceback format
python -m pytest tests/ -v --tb=short
```

### Coverage Analysis
```bash
# Run tests with coverage report
python -m pytest tests/ --cov=barrett_calculator --cov-report=term-missing

# Generate HTML coverage report
python -m pytest tests/ --cov=barrett_calculator --cov-report=html
```

### Specific Test Categories
```bash
# Run only specific test class
python -m pytest tests/test_barrett_calculator.py::TestBarrettCalculator -v

# Run only edge case tests
python -m pytest tests/test_barrett_calculator.py::TestBarrettCalculatorEdgeCases -v

# Run specific test method
python -m pytest tests/test_barrett_calculator.py::TestBarrettCalculator::test_load_patient_data_success -v
```

## Test Results

Current test metrics:
- **43 test cases** total
- **90% code coverage** of `barrett_calculator.py`
- **All tests passing** ✅

### Coverage Details
The test suite achieves 90% code coverage, missing only:
- Some specific exception handling branches
- Main execution function (`main()`)
- Some edge cases in error handling

## Mock Strategy

The tests use comprehensive mocking for:

1. **File Operations**: Temporary files and controlled I/O scenarios
2. **Playwright/Web Automation**: Mock page objects and browser interactions
3. **Excel Operations**: Controlled pandas DataFrame operations
4. **Network/External Dependencies**: Isolated unit testing

## Test Data

### Sample Patient Data
Tests use realistic patient data structures matching the actual Excel format:
```python
{
    'DoctorName': 'Dr. Smith',
    'PatientName': 'Test Patient 1',
    'PatientID': '12345',
    'LensFactor': 1.83,
    'AConstant': 119.0,
    'AxialLength_R': 23.04,
    'MeasuredK1_R': 44.75,
    'MeasuredK2_R': 44.25,
    'OpticalACD_R': 2.18,
    'Refraction_R': -0.03,
    'IOLPower': 21.5,
    'Optic': 'Biconvex',
    'Refraction': None
}
```

### Fixtures
Shared fixtures provide:
- Sample patient data (`sample_patient_data`)
- Temporary Excel files (`temp_excel_file`)
- Empty Excel files (`empty_excel_file`)
- Mock Playwright page objects (`mock_page`)
- Calculator instances (`calculator`)

## Best Practices Implemented

1. **Test Isolation**: Each test is independent with proper setup/teardown
2. **Comprehensive Mocking**: External dependencies are properly mocked
3. **Clear Naming**: Test names follow `test_<method>_<scenario>_<expected>` pattern
4. **Edge Case Coverage**: Tests include both normal and abnormal scenarios
5. **Parametrized Testing**: Multiple input scenarios tested efficiently
6. **Error Handling**: Exception paths are thoroughly tested
7. **Fixture Reuse**: Common test data and objects are shared via fixtures

## Maintenance

When modifying the `BarrettCalculator` class:

1. **Add corresponding tests** for new functionality
2. **Update existing tests** if method signatures change
3. **Maintain coverage** above 85%
4. **Run full test suite** before committing changes
5. **Update test data** if Excel format changes

## Troubleshooting

### Common Issues

1. **Import Errors**: Ensure the project root is in Python path
2. **Fixture Not Found**: Check fixture scope and location
3. **Mock Failures**: Verify mock setup matches actual method calls
4. **Coverage Gaps**: Add tests for uncovered code paths

### Debug Mode
Run tests with more verbose output:
```bash
python -m pytest tests/ -v -s --tb=long
```