import pytest
from fastapi.testclient import TestClient
from src.main import app

client = TestClient(app)

def test_transform_excel_endpoint(sample_excel_file):
    response = client.post(
        "/transform_excel",
        json={
            "file": sample_excel_file,
            "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "operations": [
                {
                    "sheet_name": "Sheet1",
                    "processing": [
                        {
                            "processing_type": "set_cells",
                            "target": {
                                "cells": {
                                    "start_cell": {
                                        "col_letter": "C",
                                        "row": 1
                                    }
                                },
                                "values": [["New Value"]]
                            }
                        }
                    ]
                }
            ]
        }
    )
    
    assert response.status_code == 200
    assert response.json()["status"] == "Success"
    assert "output" in response.json()