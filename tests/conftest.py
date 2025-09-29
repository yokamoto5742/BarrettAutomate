"""
Pytest configuration and shared fixtures for BarrettAutomate tests.
"""
import tempfile
from pathlib import Path

import pandas as pd
import pytest


@pytest.fixture
def sample_patient_data():
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
def temp_excel_file(sample_patient_data):
    """Create temporary Excel file with sample data"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        sample_patient_data.to_excel(tmp.name, index=False)
        yield tmp.name
    Path(tmp.name).unlink(missing_ok=True)


@pytest.fixture
def empty_excel_file():
    """Create temporary empty Excel file"""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        pd.DataFrame().to_excel(tmp.name, index=False)
        yield tmp.name
    Path(tmp.name).unlink(missing_ok=True)