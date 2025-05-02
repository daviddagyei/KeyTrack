import os
import sys
import pytest
import shutil
import openpyxl
from datetime import datetime
from pathlib import Path

# Add the parent directory to sys.path to import modules correctly
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from models.entities import Room, Student, KeyAction
from models.exceptions import KeyTrackException
from repository.excel_repository import ExcelRepository

@pytest.fixture
def test_excel_path():
    """Create a temporary Excel file for testing."""
    test_file = "test_key_distribution.xlsx"
    
    # Create test Excel file
    wb = openpyxl.Workbook()
    ws = wb.active
    
    # Add header
    ws.append(["Room", "Available Keys", "Collected By", "Lost Keys", "Borrowed Spare Keys", "Returned Keys"])
    
    # Add test data
    ws.append(["101", 5, "", "", "", ""])
    ws.append(["102", 3, "Student1 (2025-01-01 10:00:00)", "", "", ""])
    ws.append(["103", 0, "Student1 (2025-01-01 10:00:00), Student2 (2025-01-01 11:00:00)", 
               "Student3 (2025-01-02 09:00:00)", "", ""])
    
    wb.save(test_file)
    yield test_file
    
    # Cleanup
    if os.path.exists(test_file):
        os.remove(test_file)

@pytest.fixture
def excel_repository(test_excel_path):
    """Create a repository instance for testing."""
    return ExcelRepository(test_excel_path)

@pytest.fixture
def sample_student():
    """Create a sample student for testing."""
    return Student(id="S_1234", name="Test Student")

@pytest.fixture
def sample_room():
    """Create a sample room for testing."""
    room = Room(id="101", total_keys=5)
    return room

@pytest.fixture
def sample_key_action(sample_student):
    """Create a sample key action for testing."""
    return KeyAction(student=sample_student, timestamp=datetime.now())