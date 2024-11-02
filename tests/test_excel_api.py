import pytest
from fastapi.testclient import TestClient
from fastapi import FastAPI
import base64
from src.schemas.models import ExcelRequest
from main import app  # assuming your FastAPI app is in main.py
import io

client = TestClient(app)

def create_sample_excel_file():
    """Helper function to create a sample Excel file and return its base64 encoding"""
    import pandas as pd
    
    # Create a simple Excel file in memory
    buffer = io.BytesIO()
    df = pd.DataFrame({
        'Column1': [1, 2, 3],
        'Column2': ['A', 'B', 'C']
    })
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    
    # Convert to base64
    return base64.b64encode(buffer.read()).decode()

def test_transform_excel_success():
    """Test successful Excel transformation"""
    # Prepare test data
    sample_excel = create_sample_excel_file()
    test_request = {
        "file": sample_excel,
        "operations": [
            {
                "sheet_name": "Sheet1",
                "processing": {
                    "operation": "sort",
                    "column": "Column1"
                }
            }
        ]
    }

    # Make request to the endpoint
    response = client.post("/transform_excel", json=test_request)

    # Assert response
    assert response.status_code == 200
    assert response.json()["status"] == "Success"
    assert response.json()["error_code"] == 200
    assert "output" in response.json()
    assert response.json()["mimetype"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def test_transform_excel_invalid_file():
    """Test error handling with invalid Excel file"""
    test_request = {
        "file": "invalid_base64_content",
        "operations": []
    }

    response = client.post("/transform_excel", json=test_request)

    assert response.status_code == 400
    assert response.json()["status"] == "Error"
    assert response.json()["error_code"] == 400

def test_transform_excel_invalid_operation():
    """Test error handling with invalid operation"""
    sample_excel = create_sample_excel_file()
    test_request = {
        "file": sample_excel,
        "operations": [
            {
                "sheet_name": "NonexistentSheet",
                "processing": {
                    "operation": "invalid_operation"
                }
            }
        ]
    }

    response = client.post("/transform_excel", json=test_request)

    assert response.status_code == 400
    assert response.json()["status"] == "Error"
    assert response.json()["error_code"] == 400