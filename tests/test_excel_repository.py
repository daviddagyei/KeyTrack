import pytest
import os
import shutil
import tempfile
from datetime import datetime
from unittest.mock import patch, MagicMock
from openpyxl import Workbook

from models.entities import Room, Student, KeyAction
from models.exceptions import DataAccessException
from repository.excel_repository import ExcelRepository


class TestExcelRepository:
    """Test cases for the Excel repository implementation."""
    
    def test_initialization(self, test_excel_path):
        """Test that the repository initializes correctly."""
        repo = ExcelRepository(test_excel_path)
        assert repo.file_path == test_excel_path
        assert repo._workbook is None
        assert repo._sheet is None
    
    def test_open_workbook_success(self, test_excel_path):
        """Test opening a workbook successfully."""
        repo = ExcelRepository(test_excel_path)
        repo._open_workbook()
        assert repo._workbook is not None
        assert repo._sheet is not None
        repo._close_workbook()
    
    def test_open_workbook_file_not_found(self):
        """Test opening a non-existent workbook."""
        repo = ExcelRepository("nonexistent_file.xlsx")
        with pytest.raises(DataAccessException) as exc_info:
            repo._open_workbook()
        assert "Failed to open Excel file" in str(exc_info.value)
    
    def test_close_workbook_without_save(self, test_excel_path):
        """Test closing a workbook without saving."""
        repo = ExcelRepository(test_excel_path)
        repo._open_workbook()
        assert repo._workbook is not None
        
        repo._close_workbook(save=False)
        assert repo._workbook is None
        assert repo._sheet is None
    
    def test_close_workbook_with_save(self, test_excel_path):
        """Test closing a workbook with saving."""
        repo = ExcelRepository(test_excel_path)
        repo._open_workbook()
        assert repo._workbook is not None
        
        repo._close_workbook(save=True)
        assert repo._workbook is None
        assert repo._sheet is None
    
    def test_close_workbook_with_save_exception(self, test_excel_path):
        """Test exception during workbook save."""
        repo = ExcelRepository(test_excel_path)
        repo._open_workbook()
        
        # Make the workbook save method throw an exception
        with patch.object(repo._workbook, 'save', side_effect=IOError("Mock save error")):
            with pytest.raises(DataAccessException) as exc_info:
                repo._close_workbook(save=True)
            assert "Failed to save Excel file" in str(exc_info.value)
        
        # Even after exception, the workbook should be closed
        assert repo._workbook is None
        assert repo._sheet is None
    
    def test_parse_list_normal(self, excel_repository):
        """Test parsing a normal comma-separated list."""
        result = excel_repository._parse_list("a, b, c")
        assert result == ["a", "b", "c"]
    
    def test_parse_list_empty(self, excel_repository):
        """Test parsing an empty string."""
        result = excel_repository._parse_list("")
        assert result == []
    
    def test_parse_list_none(self, excel_repository):
        """Test parsing a None value."""
        result = excel_repository._parse_list(None)
        assert result == []
    
    def test_parse_list_single_item(self, excel_repository):
        """Test parsing a list with a single item."""
        result = excel_repository._parse_list("single")
        assert result == ["single"]
    
    def test_parse_list_with_spaces(self, excel_repository):
        """Test parsing a list with extra spaces."""
        result = excel_repository._parse_list("a,  b,c ")
        assert result == ["a", "  b", "c "]
    
    def test_generate_student_id_consistency(self, excel_repository):
        """Test that student ID generation is consistent for the same name."""
        id1 = excel_repository._generate_student_id("John")
        id2 = excel_repository._generate_student_id("John")
        assert id1 == id2
    
    def test_generate_student_id_uniqueness(self, excel_repository):
        """Test that student ID generation produces different IDs for different names."""
        id1 = excel_repository._generate_student_id("John")
        id2 = excel_repository._generate_student_id("Mary")
        assert id1 != id2
    
    def test_generate_student_id_format(self, excel_repository):
        """Test the format of generated student IDs."""
        student_id = excel_repository._generate_student_id("John")
        assert student_id.startswith("S_")
        assert student_id[2:].isdigit()
    
    def test_generate_student_id_with_special_chars(self, excel_repository):
        """Test ID generation with special characters in the name."""
        id1 = excel_repository._generate_student_id("John O'Connor")
        id2 = excel_repository._generate_student_id("John O'Connor")
        assert id1 == id2
    
    def test_load_empty_file(self):
        """Test loading an empty Excel file."""
        # Create a temporary empty workbook
        temp_path = "empty_test.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["Room", "Available Keys", "Collected By", "Lost Keys", "Borrowed Spare Keys", "Returned Keys"])
        wb.save(temp_path)
        
        try:
            repo = ExcelRepository(temp_path)
            rooms = repo.load()
            assert isinstance(rooms, dict)
            assert len(rooms) == 0
        finally:
            # Clean up
            if os.path.exists(temp_path):
                os.remove(temp_path)
    
    def test_load_rooms_from_excel(self, excel_repository):
        """Test loading room data from Excel with verification of all fields."""
        rooms = excel_repository.load()
        
        # Verify room count
        assert len(rooms) == 3
        assert "101" in rooms
        assert "102" in rooms
        assert "103" in rooms
        
        # Check room 101 - Empty room
        room_101 = rooms["101"]
        assert room_101.id == "101"
        assert room_101.total_keys == 5
        assert room_101.available_keys == 5
        assert len(room_101.collected_actions) == 0
        assert len(room_101.lost_actions) == 0
        assert len(room_101.borrowed_actions) == 0
        assert len(room_101.returned_actions) == 0
        
        # Check room 102 - Room with collected key
        room_102 = rooms["102"]
        assert room_102.id == "102"
        assert room_102.total_keys == 3
        assert len(room_102.collected_actions) == 1
        assert room_102.collected_actions[0].student.name == "Student1"
        assert room_102.collected_actions[0].timestamp == datetime(2025, 1, 1, 10, 0, 0)
        
        # Check room 103 - Room with multiple actions
        room_103 = rooms["103"]
        assert room_103.id == "103"
        assert room_103.total_keys == 0
        assert len(room_103.collected_actions) == 2
        assert room_103.collected_actions[0].student.name == "Student1"
        assert room_103.collected_actions[1].student.name == "Student2"
        assert len(room_103.lost_actions) == 1
        assert room_103.lost_actions[0].student.name == "Student3"
    
    def test_save_single_room_to_excel(self, excel_repository):
        """Test saving a single room to Excel."""
        # First load the data
        rooms = excel_repository.load()
        
        # Modify room 101
        room_101 = rooms["101"]
        student = Student(id="S_1234", name="Test Student")
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action = KeyAction(student=student, timestamp=timestamp)
        
        room_101.collected_actions.append(action)
        
        # Save the data
        assert excel_repository.save(rooms) is True
        
        # Load the data again to verify changes
        updated_rooms = excel_repository.load()
        updated_room_101 = updated_rooms["101"]
        
        assert len(updated_room_101.collected_actions) == 1
        assert updated_room_101.collected_actions[0].student.name == "Test Student"
        assert updated_room_101.collected_actions[0].timestamp == timestamp
    
    def test_save_multiple_rooms_to_excel(self, excel_repository):
        """Test saving multiple rooms to Excel."""
        # First load the data
        rooms = excel_repository.load()
        
        # Modify room 101
        room_101 = rooms["101"]
        student1 = Student(id="S_1234", name="Student One")
        timestamp1 = datetime(2025, 5, 1, 10, 30, 0)
        action1 = KeyAction(student=student1, timestamp=timestamp1)
        room_101.collected_actions.append(action1)
        
        # Modify room 102
        room_102 = rooms["102"]
        student2 = Student(id="S_5678", name="Student Two")
        timestamp2 = datetime(2025, 5, 1, 11, 30, 0)
        action2 = KeyAction(student=student2, timestamp=timestamp2)
        room_102.lost_actions.append(action2)
        
        # Save the data
        assert excel_repository.save(rooms) is True
        
        # Load the data again to verify changes
        updated_rooms = excel_repository.load()
        
        # Verify room 101
        updated_room_101 = updated_rooms["101"]
        assert len(updated_room_101.collected_actions) == 1
        assert updated_room_101.collected_actions[0].student.name == "Student One"
        
        # Verify room 102
        updated_room_102 = updated_rooms["102"]
        assert len(updated_room_102.lost_actions) == 1
        assert updated_room_102.lost_actions[0].student.name == "Student Two"
    
    def test_save_with_all_action_types(self, excel_repository):
        """Test saving a room with all types of actions."""
        # First load the data
        rooms = excel_repository.load()
        
        # Modify room 101 with all action types
        room_101 = rooms["101"]
        student = Student(id="S_1234", name="Test Student")
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        
        collect_action = KeyAction(student=student, timestamp=timestamp)
        lost_action = KeyAction(student=student, timestamp=timestamp)
        borrowed_action = KeyAction(student=student, timestamp=timestamp)
        returned_action = KeyAction(student=student, timestamp=timestamp)
        
        room_101.collected_actions.append(collect_action)
        room_101.lost_actions.append(lost_action)
        room_101.borrowed_actions.append(borrowed_action)
        room_101.returned_actions.append(returned_action)
        
        # Save the data
        assert excel_repository.save(rooms) is True
        
        # Load the data again to verify changes
        updated_rooms = excel_repository.load()
        updated_room_101 = updated_rooms["101"]
        
        assert len(updated_room_101.collected_actions) == 1
        assert updated_room_101.collected_actions[0].student.name == "Test Student"
        
        assert len(updated_room_101.lost_actions) == 1
        assert updated_room_101.lost_actions[0].student.name == "Test Student"
        
        assert len(updated_room_101.borrowed_actions) == 1
        assert updated_room_101.borrowed_actions[0].student.name == "Test Student"
        
        assert len(updated_room_101.returned_actions) == 1
        assert updated_room_101.returned_actions[0].student.name == "Test Student"
    
    def test_load_nonexistent_file(self):
        """Test loading a non-existent Excel file."""
        repo = ExcelRepository("nonexistent_file.xlsx")
        with pytest.raises(DataAccessException):
            repo.load()
    
    def test_save_without_permission(self, test_excel_path):
        """Test saving to a file without permission."""
        # Make file read-only if possible (might not work in all environments)
        try:
            os.chmod(test_excel_path, 0o444)  # Read-only
            
            repo = ExcelRepository(test_excel_path)
            rooms = repo.load()
            
            with pytest.raises(DataAccessException):
                repo.save(rooms)
        finally:
            # Restore permissions
            os.chmod(test_excel_path, 0o644)
    
    def test_parse_list_method(self, excel_repository):
        """Test the _parse_list private method."""
        assert excel_repository._parse_list("") == []
        assert excel_repository._parse_list(None) == []
        assert excel_repository._parse_list("a, b, c") == ["a", "b", "c"]
    
    def test_generate_student_id(self, excel_repository):
        """Test the _generate_student_id private method."""
        id1 = excel_repository._generate_student_id("John")
        id2 = excel_repository._generate_student_id("Mary")
        id3 = excel_repository._generate_student_id("John")
        
        assert id1.startswith("S_")
        assert id2.startswith("S_")
        assert id1 != id2  # Different names should generate different IDs
        assert id1 == id3  # Same name should generate same ID