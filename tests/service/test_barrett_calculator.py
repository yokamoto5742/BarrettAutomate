from pathlib import Path
from unittest.mock import MagicMock

import pandas as pd
import pytest
from pytest_mock import MockerFixture

from service.barrett_calculator import BarrettCalculator


@pytest.fixture
def excel_path(tmp_path: Path) -> Path:
    path = tmp_path / "Barrettdata.xlsx"
    pd.DataFrame({
        "PatientName": ["患者A", "患者B", "患者C"],
        "IOL_power": [20.0, 21.0, 22.0],
    }).to_excel(path, index=False)
    return path


@pytest.fixture
def calculator(excel_path: Path, mocker: MockerFixture) -> BarrettCalculator:
    mocker.patch("service.barrett_calculator.sync_playwright", return_value=MagicMock())
    mocker.patch("service.barrett_calculator.time.sleep")
    mocker.patch("service.barrett_calculator.win32api.MessageBox")
    return BarrettCalculator(str(excel_path), headless=True)


def find_result_file(excel_path: Path) -> Path:
    results = list(excel_path.parent.glob("Barrettdata_results_*.xlsx"))
    assert len(results) == 1
    return results[0]


def test_save_patient_data_writes_file(calculator: BarrettCalculator, tmp_path: Path) -> None:
    output_path = tmp_path / "out.xlsx"
    df = pd.DataFrame({"Refraction": [-0.5]})

    assert calculator.save_patient_data(df, output_path) is True
    assert pd.read_excel(output_path)["Refraction"].tolist() == [-0.5]


def test_save_patient_data_returns_false_on_failure(
        calculator: BarrettCalculator, tmp_path: Path, mocker: MockerFixture) -> None:
    mocker.patch.object(pd.DataFrame, "to_excel", side_effect=PermissionError("locked"))

    assert calculator.save_patient_data(pd.DataFrame(), tmp_path / "out.xlsx") is False


def test_process_all_patients_saves_after_each_patient(
        calculator: BarrettCalculator, mocker: MockerFixture) -> None:
    mocker.patch.object(calculator, "input_patient_data", return_value=True)
    mocker.patch.object(calculator, "calculate_and_get_result", side_effect=[-0.1, -0.2, -0.3])
    save_spy = mocker.spy(calculator, "save_patient_data")

    calculator.process_all_patients()

    assert save_spy.call_count == 3
    output_paths = {call.args[1] for call in save_spy.call_args_list}
    assert len(output_paths) == 1


def test_results_remain_when_process_is_interrupted(
        calculator: BarrettCalculator, excel_path: Path, mocker: MockerFixture) -> None:
    mocker.patch.object(calculator, "input_patient_data", return_value=True)
    mocker.patch.object(
        calculator, "calculate_and_get_result", side_effect=[-0.1, KeyboardInterrupt()])

    with pytest.raises(KeyboardInterrupt):
        calculator.process_all_patients()

    saved = pd.read_excel(find_result_file(excel_path))
    assert saved.loc[0, "Refraction"] == -0.1
    assert pd.isna(saved.loc[1, "Refraction"])


def test_processing_continues_when_save_fails(
        calculator: BarrettCalculator, mocker: MockerFixture) -> None:
    mocker.patch.object(calculator, "input_patient_data", return_value=True)
    calculate_mock = mocker.patch.object(
        calculator, "calculate_and_get_result", side_effect=[-0.1, -0.2, -0.3])
    mocker.patch.object(calculator, "save_patient_data", return_value=False)

    calculator.process_all_patients()

    assert calculate_mock.call_count == 3


def test_processing_continues_when_calculation_fails_after_success(
        calculator: BarrettCalculator, excel_path: Path, mocker: MockerFixture) -> None:
    mocker.patch.object(calculator, "input_patient_data", return_value=True)
    mocker.patch.object(calculator, "calculate_and_get_result", side_effect=[-0.1, None, -0.3])

    calculator.process_all_patients()

    saved = pd.read_excel(find_result_file(excel_path))
    assert saved["Refraction"].tolist() == [-0.1, "計算エラー", -0.3]
