import logging
import re
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, Mock, patch

import pandas as pd
import pytest
from playwright.sync_api import Page

from barrett_calculator import BarrettCalculator


class TestBarrettCalculator:
    """Test suite for BarrettCalculator class"""

    @pytest.fixture
    def sample_excel_data(self):
        """Sample patient data for testing"""
        return pd.DataFrame([
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
            },
            {
                'DoctorName': 'Dr. Jones',
                'PatientName': 'Test Patient 2',
                'PatientID': '67890',
                'LensFactor': 1.85,
                'AConstant': 118.5,
                'AxialLength_R': 22.85,
                'MeasuredK1_R': 45.00,
                'MeasuredK2_R': 44.50,
                'OpticalACD_R': 2.25,
                'Refraction_R': 0.12,
                'IOLPower': 22.0,
                'Optic': 'Biconvex',
                'Refraction': None
            }
        ])

    @pytest.fixture
    def temp_excel_file(self, sample_excel_data):
        """Create temporary Excel file for testing"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            sample_excel_data.to_excel(tmp.name, index=False)
            yield tmp.name
        Path(tmp.name).unlink(missing_ok=True)

    @pytest.fixture
    def calculator(self, temp_excel_file):
        """Create BarrettCalculator instance for testing"""
        return BarrettCalculator(temp_excel_file, headless=True)

    @pytest.fixture
    def mock_page(self):
        """Mock Playwright page object"""
        page = Mock(spec=Page)
        page.wait_for_load_state = Mock()
        page.locator = Mock()
        page.goto = Mock()
        page.wait_for_timeout = Mock()
        page.content = Mock(return_value="<html><body>Mock content</body></html>")
        return page

    def test_init_with_valid_file(self, temp_excel_file):
        """Test BarrettCalculator initialization with valid file"""
        calculator = BarrettCalculator(temp_excel_file, headless=True)

        assert calculator.excel_file_path == Path(temp_excel_file)
        assert calculator.results_file_path.name == 'APACdata_results.xlsx'
        assert calculator.headless is True
        assert calculator.url == "https://calc.apacrs.org/barrett_universal2105/"
        assert calculator.logger is not None

    def test_init_with_custom_headless(self, temp_excel_file):
        """Test BarrettCalculator initialization with custom headless setting"""
        calculator = BarrettCalculator(temp_excel_file, headless=False)
        assert calculator.headless is False

    def test_load_patient_data_success(self, calculator, sample_excel_data):
        """Test successful patient data loading"""
        df = calculator.load_patient_data()

        assert isinstance(df, pd.DataFrame)
        assert len(df) == 2
        assert list(df.columns) == list(sample_excel_data.columns)
        assert df.iloc[0]['PatientName'] == 'Test Patient 1'
        assert df.iloc[1]['PatientName'] == 'Test Patient 2'

    def test_load_patient_data_file_not_found(self):
        """Test load_patient_data with non-existent file"""
        calculator = BarrettCalculator("nonexistent_file.xlsx", headless=True)

        with pytest.raises(FileNotFoundError) as exc_info:
            calculator.load_patient_data()

        assert "ファイルが見つかりません" in str(exc_info.value)

    @patch('pandas.read_excel')
    def test_load_patient_data_read_error(self, mock_read_excel, calculator):
        """Test load_patient_data with pandas read error"""
        mock_read_excel.side_effect = Exception("Read error")

        with pytest.raises(Exception) as exc_info:
            calculator.load_patient_data()

        assert "Read error" in str(exc_info.value)

    def test_save_patient_data_success(self, calculator, sample_excel_data, tmp_path):
        """Test successful patient data saving"""
        # Set a custom results path for testing
        calculator.results_file_path = tmp_path / "test_results.xlsx"

        calculator.save_patient_data(sample_excel_data)

        assert calculator.results_file_path.exists()

        # Verify the saved data
        saved_df = pd.read_excel(calculator.results_file_path)
        assert len(saved_df) == 2
        assert saved_df.iloc[0]['PatientName'] == 'Test Patient 1'

    def test_save_patient_data_with_backup(self, calculator, sample_excel_data, tmp_path):
        """Test saving with existing file creates backup"""
        calculator.results_file_path = tmp_path / "test_results.xlsx"

        # Create existing file
        sample_excel_data.to_excel(calculator.results_file_path, index=False)

        # Save new data
        modified_data = sample_excel_data.copy()
        modified_data.loc[0, 'PatientName'] = 'Modified Patient'

        calculator.save_patient_data(modified_data)

        # Check backup was created - fix the backup naming logic
        backup_path = tmp_path / "test_results_backup.xlsx"
        assert backup_path.exists()

        # Verify original data in backup
        backup_df = pd.read_excel(backup_path)
        assert backup_df.iloc[0]['PatientName'] == 'Test Patient 1'

        # Verify new data in main file
        main_df = pd.read_excel(calculator.results_file_path)
        assert main_df.iloc[0]['PatientName'] == 'Modified Patient'

    @patch('pandas.DataFrame.to_excel')
    def test_save_patient_data_write_error(self, mock_to_excel, calculator, sample_excel_data):
        """Test save_patient_data with write error"""
        mock_to_excel.side_effect = Exception("Write error")

        with pytest.raises(Exception) as exc_info:
            calculator.save_patient_data(sample_excel_data)

        assert "Write error" in str(exc_info.value)

    def test_input_patient_data_success(self, calculator, mock_page, sample_excel_data):
        """Test successful patient data input"""
        # Setup mock text inputs
        mock_inputs = [Mock() for _ in range(15)]  # Enough for all fields
        for mock_input in mock_inputs:
            mock_input.clear = Mock()
            mock_input.fill = Mock()

        mock_page.locator.return_value.all.return_value = mock_inputs

        patient_row = sample_excel_data.iloc[0]
        result = calculator.input_patient_data(mock_page, patient_row)

        assert result is True
        mock_page.wait_for_load_state.assert_called_with('networkidle')
        mock_page.locator.assert_called_with('input[type="text"]')

        # Verify some inputs were called
        assert any(mock_input.fill.called for mock_input in mock_inputs[:5])

    def test_input_patient_data_with_missing_fields(self, calculator, mock_page, sample_excel_data):
        """Test input_patient_data with missing optional fields"""
        # Setup mock text inputs
        mock_inputs = [Mock() for _ in range(10)]
        for mock_input in mock_inputs:
            mock_input.clear = Mock()
            mock_input.fill = Mock()

        mock_page.locator.return_value.all.return_value = mock_inputs

        # Create patient row with missing data
        patient_row = sample_excel_data.iloc[0].copy()
        patient_row['DoctorName'] = None
        patient_row['OpticalACD_R'] = ''

        result = calculator.input_patient_data(mock_page, patient_row)

        assert result is True

    def test_input_patient_data_exception(self, calculator, mock_page, sample_excel_data):
        """Test input_patient_data with exception during input"""
        mock_page.wait_for_load_state.side_effect = Exception("Network error")

        patient_row = sample_excel_data.iloc[0]
        result = calculator.input_patient_data(mock_page, patient_row)

        assert result is False

    def test_calculate_and_get_result_success(self, calculator, mock_page):
        """Test successful calculation and result extraction"""
        # Setup mocks for calculate button
        mock_calculate_btn = Mock()
        mock_calculate_btn.is_visible.return_value = True
        mock_calculate_btn.click = Mock()

        # Setup mocks for universal formula tab
        mock_universal_btn = Mock()
        mock_universal_btn.is_visible.return_value = True
        mock_universal_btn.click = Mock()

        # Setup the locator calls properly
        def locator_side_effect(selector):
            if 'Calculate' in selector:
                mock_locator = Mock()
                mock_locator.first = mock_calculate_btn
                return mock_locator
            elif 'Universal Formula' in selector:
                mock_locator = Mock()
                mock_locator.first = mock_universal_btn
                return mock_locator
            return Mock()

        mock_page.locator.side_effect = locator_side_effect

        # Mock the result extraction
        with patch.object(calculator, '_extract_refraction_from_table', return_value=1.25):
            result = calculator.calculate_and_get_result(mock_page, 21.5)

        assert result == 1.25
        mock_calculate_btn.click.assert_called_once()
        mock_universal_btn.click.assert_called_once()

    def test_calculate_and_get_result_calculate_button_not_found(self, calculator, mock_page):
        """Test calculate_and_get_result when calculate button is not found"""
        mock_calculate_btn = Mock()
        mock_calculate_btn.is_visible.return_value = False

        mock_locator = Mock()
        mock_locator.first = mock_calculate_btn
        mock_page.locator.return_value = mock_locator

        result = calculator.calculate_and_get_result(mock_page, 21.5)

        assert result is None

    def test_calculate_and_get_result_universal_tab_not_found(self, calculator, mock_page):
        """Test calculate_and_get_result when universal formula tab is not found"""
        # Setup calculate button
        mock_calculate_btn = Mock()
        mock_calculate_btn.is_visible.return_value = True
        mock_calculate_btn.click = Mock()

        # Setup universal tab (not visible)
        mock_universal_btn = Mock()
        mock_universal_btn.is_visible.return_value = False

        # Setup the locator calls properly
        def locator_side_effect(selector):
            if 'Calculate' in selector:
                mock_locator = Mock()
                mock_locator.first = mock_calculate_btn
                return mock_locator
            elif 'Universal Formula' in selector:
                mock_locator = Mock()
                mock_locator.first = mock_universal_btn
                return mock_locator
            return Mock()

        mock_page.locator.side_effect = locator_side_effect

        result = calculator.calculate_and_get_result(mock_page, 21.5)

        assert result is None

    def test_calculate_and_get_result_exception(self, calculator, mock_page):
        """Test calculate_and_get_result with exception"""
        mock_page.locator.side_effect = Exception("Page error")

        result = calculator.calculate_and_get_result(mock_page, 21.5)

        assert result is None

    def test_extract_refraction_from_table_success(self, calculator, mock_page):
        """Test successful refraction extraction from table"""
        # Create mock table rows and cells
        mock_cell1 = Mock()
        mock_cell1.text_content.return_value = "21.5"
        mock_cell2 = Mock()
        mock_cell2.text_content.return_value = "Biconvex"
        mock_cell3 = Mock()
        mock_cell3.text_content.return_value = "1.25"

        mock_row = Mock()
        mock_row.locator.return_value.all.return_value = [mock_cell1, mock_cell2, mock_cell3]

        mock_page.locator.return_value.all.return_value = [mock_row]

        result = calculator._extract_refraction_from_table(mock_page, 21.5)

        assert result == 1.25

    def test_extract_refraction_from_table_no_match(self, calculator, mock_page):
        """Test refraction extraction with no matching IOL power"""
        # Create mock table row with different IOL power
        mock_cell1 = Mock()
        mock_cell1.text_content.return_value = "20.0"  # Different IOL power
        mock_cell2 = Mock()
        mock_cell2.text_content.return_value = "Biconvex"
        mock_cell3 = Mock()
        mock_cell3.text_content.return_value = "1.25"

        mock_row = Mock()
        mock_row.locator.return_value.all.return_value = [mock_cell1, mock_cell2, mock_cell3]

        mock_page.locator.return_value.all.return_value = [mock_row]

        with patch.object(calculator, '_extract_refraction_alternative', return_value=None):
            result = calculator._extract_refraction_from_table(mock_page, 21.5)

        assert result is None

    def test_extract_refraction_from_table_invalid_data(self, calculator, mock_page):
        """Test refraction extraction with invalid table data"""
        # Create mock table row with invalid data
        mock_cell1 = Mock()
        mock_cell1.text_content.return_value = "invalid"

        mock_row = Mock()
        mock_row.locator.return_value.all.return_value = [mock_cell1]

        mock_page.locator.return_value.all.return_value = [mock_row]

        with patch.object(calculator, '_extract_refraction_alternative', return_value=None):
            result = calculator._extract_refraction_from_table(mock_page, 21.5)

        assert result is None

    def test_extract_refraction_from_table_exception(self, calculator, mock_page):
        """Test refraction extraction with exception"""
        mock_page.locator.side_effect = Exception("Table error")

        result = calculator._extract_refraction_from_table(mock_page, 21.5)

        assert result is None

    def test_extract_refraction_alternative_success(self, calculator, mock_page):
        """Test alternative refraction extraction method"""
        mock_page.content.return_value = """
        <table>
            <tr><td>21.5</td><td>Biconvex</td><td>1.25</td></tr>
        </table>
        """
        mock_page.locator.return_value.all.return_value = []

        result = calculator._extract_refraction_alternative(mock_page, 21.5)

        assert result == 1.25

    def test_extract_refraction_alternative_highlighted_row(self, calculator, mock_page):
        """Test alternative extraction with highlighted row"""
        mock_page.content.return_value = "<html><body>No pattern match</body></html>"

        # Mock highlighted row
        mock_highlighted_row = Mock()
        mock_highlighted_row.text_content.return_value = "21.5 Biconvex 1.75"
        mock_page.locator.return_value.all.return_value = [mock_highlighted_row]

        result = calculator._extract_refraction_alternative(mock_page, 21.5)

        assert result == 1.75

    def test_extract_refraction_alternative_no_match(self, calculator, mock_page):
        """Test alternative extraction with no matches"""
        mock_page.content.return_value = "<html><body>No relevant data</body></html>"
        mock_page.locator.return_value.all.return_value = []

        result = calculator._extract_refraction_alternative(mock_page, 21.5)

        assert result is None

    def test_extract_refraction_alternative_exception(self, calculator, mock_page):
        """Test alternative extraction with exception"""
        mock_page.content.side_effect = Exception("Content error")

        result = calculator._extract_refraction_alternative(mock_page, 21.5)

        assert result is None

    @patch('barrett_calculator.sync_playwright')
    def test_process_all_patients_success(self, mock_playwright, calculator, sample_excel_data):
        """Test successful processing of all patients"""
        # Setup playwright mocks
        mock_browser = Mock()
        mock_context = Mock()
        mock_page = Mock()

        mock_context.new_page.return_value = mock_page
        mock_browser.new_context.return_value = mock_context
        mock_playwright.return_value.__enter__.return_value.chromium.launch.return_value = mock_browser

        # Mock successful data input and calculation
        with patch.object(calculator, 'load_patient_data', return_value=sample_excel_data), \
             patch.object(calculator, 'input_patient_data', return_value=True), \
             patch.object(calculator, 'calculate_and_get_result', return_value=1.25), \
             patch.object(calculator, 'save_patient_data') as mock_save:

            calculator.process_all_patients()

            mock_save.assert_called_once()
            # Verify the dataframe was updated with results
            saved_df = mock_save.call_args[0][0]
            assert saved_df.loc[0, 'Refraction'] == 1.25
            assert saved_df.loc[1, 'Refraction'] == 1.25

    @patch('barrett_calculator.sync_playwright')
    def test_process_all_patients_input_error(self, mock_playwright, calculator, sample_excel_data):
        """Test processing with input error"""
        # Setup playwright mocks
        mock_browser = Mock()
        mock_context = Mock()
        mock_page = Mock()

        mock_context.new_page.return_value = mock_page
        mock_browser.new_context.return_value = mock_context
        mock_playwright.return_value.__enter__.return_value.chromium.launch.return_value = mock_browser

        # Mock failed data input
        with patch.object(calculator, 'load_patient_data', return_value=sample_excel_data), \
             patch.object(calculator, 'input_patient_data', return_value=False), \
             patch.object(calculator, 'save_patient_data') as mock_save:

            calculator.process_all_patients()

            mock_save.assert_called_once()
            saved_df = mock_save.call_args[0][0]
            assert saved_df.loc[0, 'Refraction'] == "入力エラー"

    @patch('barrett_calculator.sync_playwright')
    def test_process_all_patients_calculation_error(self, mock_playwright, calculator, sample_excel_data):
        """Test processing with calculation error"""
        # Setup playwright mocks
        mock_browser = Mock()
        mock_context = Mock()
        mock_page = Mock()

        mock_context.new_page.return_value = mock_page
        mock_browser.new_context.return_value = mock_context
        mock_playwright.return_value.__enter__.return_value.chromium.launch.return_value = mock_browser

        # Mock successful input but failed calculation
        with patch.object(calculator, 'load_patient_data', return_value=sample_excel_data), \
             patch.object(calculator, 'input_patient_data', return_value=True), \
             patch.object(calculator, 'calculate_and_get_result', return_value=None), \
             patch.object(calculator, 'save_patient_data') as mock_save:

            calculator.process_all_patients()

            mock_save.assert_called_once()
            saved_df = mock_save.call_args[0][0]
            assert saved_df.loc[0, 'Refraction'] == "計算エラー"

    @patch('barrett_calculator.sync_playwright')
    def test_process_all_patients_exception(self, mock_playwright, calculator, sample_excel_data):
        """Test processing with patient-level exception"""
        # Setup playwright mocks
        mock_browser = Mock()
        mock_context = Mock()
        mock_page = Mock()

        mock_context.new_page.return_value = mock_page
        mock_browser.new_context.return_value = mock_context
        mock_playwright.return_value.__enter__.return_value.chromium.launch.return_value = mock_browser

        # Mock exception during processing
        with patch.object(calculator, 'load_patient_data', return_value=sample_excel_data), \
             patch.object(calculator, 'input_patient_data', side_effect=Exception("Processing error")), \
             patch.object(calculator, 'save_patient_data') as mock_save:

            calculator.process_all_patients()

            mock_save.assert_called_once()
            saved_df = mock_save.call_args[0][0]
            assert "エラー: Processing error" in saved_df.loc[0, 'Refraction']

    def test_process_all_patients_load_error(self, calculator):
        """Test processing with data loading error"""
        with patch.object(calculator, 'load_patient_data', side_effect=Exception("Load error")):
            with pytest.raises(Exception) as exc_info:
                calculator.process_all_patients()

            assert "Load error" in str(exc_info.value)


class TestBarrettCalculatorEdgeCases:
    """Test edge cases and boundary conditions"""

    @pytest.fixture
    def temp_excel_file(self):
        """Create temporary Excel file for testing"""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            pd.DataFrame().to_excel(tmp.name, index=False)
            yield tmp.name
        Path(tmp.name).unlink(missing_ok=True)

    @pytest.fixture
    def calculator(self, temp_excel_file):
        """Create BarrettCalculator instance for testing"""
        return BarrettCalculator(temp_excel_file, headless=True)

    @pytest.fixture
    def mock_page(self):
        """Mock Playwright page object"""
        page = Mock(spec=Page)
        page.wait_for_load_state = Mock()
        page.locator = Mock()
        page.goto = Mock()
        page.wait_for_timeout = Mock()
        page.content = Mock(return_value="<html><body>Mock content</body></html>")
        return page

    def test_empty_excel_file(self, tmp_path):
        """Test with empty Excel file"""
        empty_file = tmp_path / "empty.xlsx"
        pd.DataFrame().to_excel(empty_file, index=False)

        calculator = BarrettCalculator(str(empty_file), headless=True)
        df = calculator.load_patient_data()

        assert len(df) == 0

    def test_patient_data_with_nan_values(self, tmp_path):
        """Test patient data with NaN values"""
        data_with_nan = pd.DataFrame([{
            'DoctorName': 'Dr. Test',
            'PatientName': 'Patient Test',
            'PatientID': None,
            'LensFactor': float('nan'),
            'AConstant': 119.0,
            'AxialLength_R': 23.04,
            'MeasuredK1_R': None,
            'MeasuredK2_R': 44.25,
            'OpticalACD_R': 2.18,
            'Refraction_R': -0.03,
            'IOLPower': 21.5,
            'Optic': 'Biconvex',
            'Refraction': None
        }])

        excel_file = tmp_path / "test_nan.xlsx"
        data_with_nan.to_excel(excel_file, index=False)

        calculator = BarrettCalculator(str(excel_file), headless=True)
        df = calculator.load_patient_data()

        assert len(df) == 1
        assert pd.isna(df.iloc[0]['PatientID'])

    def test_negative_refraction_values(self, calculator, mock_page):
        """Test extraction of negative refraction values"""
        # Test negative value extraction
        mock_cell1 = Mock()
        mock_cell1.text_content.return_value = "21.5"
        mock_cell2 = Mock()
        mock_cell2.text_content.return_value = "Biconvex"
        mock_cell3 = Mock()
        mock_cell3.text_content.return_value = "-2.75"

        mock_row = Mock()
        mock_row.locator.return_value.all.return_value = [mock_cell1, mock_cell2, mock_cell3]

        mock_page.locator.return_value.all.return_value = [mock_row]

        result = calculator._extract_refraction_from_table(mock_page, 21.5)

        assert result == -2.75

    def test_iol_power_tolerance(self, calculator, mock_page):
        """Test IOL power matching with tolerance"""
        # Test that IOL power 21.45 matches target 21.5 (within 0.1 tolerance)
        mock_cell1 = Mock()
        mock_cell1.text_content.return_value = "21.45"  # Within tolerance
        mock_cell2 = Mock()
        mock_cell2.text_content.return_value = "Biconvex"
        mock_cell3 = Mock()
        mock_cell3.text_content.return_value = "1.25"

        mock_row = Mock()
        mock_row.locator.return_value.all.return_value = [mock_cell1, mock_cell2, mock_cell3]

        mock_page.locator.return_value.all.return_value = [mock_row]

        result = calculator._extract_refraction_from_table(mock_page, 21.5)

        assert result == 1.25

    def test_malformed_table_data(self, calculator, mock_page):
        """Test handling of malformed table data"""
        # Test table row with insufficient cells
        mock_row = Mock()
        mock_row.locator.return_value.all.return_value = [Mock()]  # Only 1 cell

        mock_page.locator.return_value.all.return_value = [mock_row]

        with patch.object(calculator, '_extract_refraction_alternative', return_value=None):
            result = calculator._extract_refraction_from_table(mock_page, 21.5)

        assert result is None


@pytest.mark.parametrize("iol_power,expected_match", [
    (21.0, True),
    (21.05, True),  # Within tolerance
    (21.15, False),  # Outside tolerance
    (20.85, False),  # Outside tolerance
])
def test_iol_power_matching_parametrized(iol_power, expected_match):
    """Parametrized test for IOL power matching logic"""
    target_iol_power = 21.0
    tolerance = 0.1

    actual_match = abs(iol_power - target_iol_power) < tolerance
    assert actual_match == expected_match


@pytest.mark.parametrize("refraction_text,expected_value", [
    ("1.25", 1.25),
    ("-2.75", -2.75),
    ("0.00", 0.0),
    ("+1.50", 1.50),
    ("invalid", None),
    ("", None),
])
def test_refraction_value_extraction(refraction_text, expected_value):
    """Parametrized test for refraction value extraction"""
    if expected_value is None:
        match = re.search(r'(-?\d+\.?\d*)', refraction_text)
        assert match is None
    else:
        match = re.search(r'(-?\d+\.?\d*)', refraction_text)
        assert match is not None
        assert float(match.group(1)) == expected_value