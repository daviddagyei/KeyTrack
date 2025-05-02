import pytest
from datetime import datetime
from unittest.mock import MagicMock, patch, call

from models.entities import Room, Student, KeyAction
from models.exceptions import RoomNotFoundException, NoKeysAvailableException, KeyNotCollectedException
from services.key_management_service import DefaultKeyManagementService


class TestKeyManagementService:
    """Test cases for the Key Management Service."""
    
    @pytest.fixture
    def mock_repository(self):
        """Create a mock repository for testing."""
        repository = MagicMock()
        
        # Set up test data
        room101 = Room(id="101", total_keys=2)
        room102 = Room(id="102", total_keys=0)  # No keys available
        room103 = Room(id="103", total_keys=1)
        
        # Add a key collection to room 103
        student = Student(id="S_1234", name="John Doe")
        action = KeyAction(student=student, timestamp=datetime(2025, 5, 1, 10, 30, 0))
        room103.collected_actions.append(action)
        
        # Mock the load method to return these rooms
        repository.load.return_value = {
            "101": room101,
            "102": room102,
            "103": room103
        }
        
        # Mock the save method to return True
        repository.save.return_value = True
        
        return repository
    
    @pytest.fixture
    def service(self, mock_repository):
        """Create a service instance for testing."""
        return DefaultKeyManagementService(repository=mock_repository)
    
    def test_initialization(self, mock_repository):
        """Test that the service initializes correctly."""
        service = DefaultKeyManagementService(repository=mock_repository)
        assert service.repository == mock_repository
        assert service._rooms is None
    
    def test_load_data_initial(self, service, mock_repository):
        """Test loading data for the first time."""
        rooms = service._load_data()
        
        assert rooms is not None
        assert len(rooms) == 3
        assert service._rooms is not None  # Data should be cached
        assert mock_repository.load.call_count == 1  # Should call repository's load method
    
    def test_load_data_cached(self, service, mock_repository):
        """Test that data is loaded only once and then cached."""
        # First call should load from repository
        rooms1 = service._load_data()
        assert mock_repository.load.call_count == 1
        
        # Subsequent calls should use cached data
        rooms2 = service._load_data()
        rooms3 = service._load_data()
        
        assert mock_repository.load.call_count == 1  # Still only one call
        assert rooms1 is rooms2  # Same object reference (cached)
        assert rooms2 is rooms3  # Same object reference (cached)
    
    def test_save_data_success(self, service, mock_repository):
        """Test saving data successfully."""
        # First load the data to cache it
        service._load_data()
        
        # Then save it
        result = service._save_data()
        
        assert result is True
        mock_repository.save.assert_called_once_with(service._rooms)
    
    def test_save_data_without_loading_first(self, service, mock_repository):
        """Test saving data without loading first."""
        result = service._save_data()
        
        assert result is False
        mock_repository.save.assert_not_called()
    
    def test_get_room_success(self, service):
        """Test getting a room that exists."""
        room = service._get_room("101")
        
        assert room is not None
        assert room.id == "101"
        assert room.total_keys == 2
    
    def test_get_room_not_found(self, service):
        """Test getting a room that doesn't exist."""
        with pytest.raises(RoomNotFoundException) as exc_info:
            service._get_room("999")
        
        assert exc_info.value.room_id == "999"
    
    def test_create_key_action(self, service):
        """Test creating a key action."""
        action = service._create_key_action("John Doe")
        
        assert action is not None
        assert action.student is not None
        assert action.student.name == "John Doe"
        assert action.student.id.startswith("S_")
        assert isinstance(action.timestamp, datetime)
    
    def test_collect_key_success(self, service):
        """Test collecting a key successfully."""
        result = service.collect_key("101", "John Doe")
        
        assert result is True
        
        # Verify the room was updated
        room = service._get_room("101")
        assert len(room.collected_actions) == 1
        assert room.collected_actions[0].student.name == "John Doe"
    
    def test_collect_key_already_has_key(self, service):
        """Test collecting a key when student already has one (should still work)."""
        # First collection
        service.collect_key("101", "John Doe")
        
        # Second collection (should still work)
        result = service.collect_key("101", "John Doe")
        
        assert result is True
        
        # Verify the room was updated with two collections
        room = service._get_room("101")
        assert len(room.collected_actions) == 2
        assert room.collected_actions[0].student.name == "John Doe"
        assert room.collected_actions[1].student.name == "John Doe"
    
    def test_collect_key_no_keys_available(self, service):
        """Test collecting a key when none are available."""
        with pytest.raises(NoKeysAvailableException) as exc_info:
            service.collect_key("102", "John Doe")
        
        assert exc_info.value.room_id == "102"
        
        # Verify no changes to the room
        room = service._get_room("102")
        assert len(room.collected_actions) == 0
    
    def test_collect_key_room_not_found(self, service):
        """Test collecting a key for a non-existent room."""
        with pytest.raises(RoomNotFoundException) as exc_info:
            service.collect_key("999", "John Doe")
        
        assert exc_info.value.room_id == "999"
    
    def test_return_key_success(self, service):
        """Test returning a key successfully."""
        # First collect a key
        service.collect_key("101", "John Doe")
        
        # Then return it
        result = service.return_key("101", "John Doe")
        
        assert result is True
        
        # Verify the room was updated
        room = service._get_room("101")
        assert len(room.collected_actions) == 1
        assert len(room.returned_actions) == 1
        assert room.returned_actions[0].student.name == "John Doe"
    
    def test_return_key_not_collected(self, service):
        """Test returning a key that wasn't collected."""
        with pytest.raises(KeyNotCollectedException) as exc_info:
            service.return_key("101", "Jane Smith")
        
        assert exc_info.value.room_id == "101"
        assert exc_info.value.student_name == "Jane Smith"
        
        # Verify no changes to the room
        room = service._get_room("101")
        assert len(room.returned_actions) == 0
    
    def test_return_key_already_returned(self, service):
        """Test returning a key that was already returned."""
        # First collect and return a key
        service.collect_key("101", "John Doe")
        service.return_key("101", "John Doe")
        
        # Attempt to return it again
        with pytest.raises(KeyNotCollectedException) as exc_info:
            service.return_key("101", "John Doe")
        
        assert exc_info.value.room_id == "101"
        assert exc_info.value.student_name == "John Doe"
    
    def test_return_key_room_not_found(self, service):
        """Test returning a key for a non-existent room."""
        with pytest.raises(RoomNotFoundException) as exc_info:
            service.return_key("999", "John Doe")
        
        assert exc_info.value.room_id == "999"
    
    def test_report_lost_key_success(self, service):
        """Test reporting a lost key successfully."""
        # Room 103 already has a collected key for John Doe
        result = service.report_lost_key("103", "John Doe")
        
        assert result is True
        
        # Verify the room was updated
        room = service._get_room("103")
        assert len(room.lost_actions) == 1
        assert room.lost_actions[0].student.name == "John Doe"
    
    def test_report_lost_key_not_collected(self, service):
        """Test reporting a lost key that wasn't collected."""
        with pytest.raises(KeyNotCollectedException) as exc_info:
            service.report_lost_key("101", "John Doe")
        
        assert exc_info.value.room_id == "101"
        assert exc_info.value.student_name == "John Doe"
        
        # Verify no changes to the room
        room = service._get_room("101")
        assert len(room.lost_actions) == 0
    
    def test_report_lost_key_room_not_found(self, service):
        """Test reporting a lost key for a non-existent room."""
        with pytest.raises(RoomNotFoundException) as exc_info:
            service.report_lost_key("999", "John Doe")
        
        assert exc_info.value.room_id == "999"
    
    def test_borrow_spare_key_success(self, service):
        """Test borrowing a spare key successfully."""
        result = service.borrow_spare_key("101", "Jane Smith")
        
        assert result is True
        
        # Verify the room was updated
        room = service._get_room("101")
        assert len(room.borrowed_actions) == 1
        assert room.borrowed_actions[0].student.name == "Jane Smith"
    
    def test_borrow_spare_key_no_keys_available(self, service):
        """Test borrowing a spare key when none are available."""
        with pytest.raises(NoKeysAvailableException) as exc_info:
            service.borrow_spare_key("102", "Jane Smith")
        
        assert exc_info.value.room_id == "102"
        
        # Verify no changes to the room
        room = service._get_room("102")
        assert len(room.borrowed_actions) == 0
    
    def test_borrow_spare_key_room_not_found(self, service):
        """Test borrowing a spare key for a non-existent room."""
        with pytest.raises(RoomNotFoundException) as exc_info:
            service.borrow_spare_key("999", "Jane Smith")
        
        assert exc_info.value.room_id == "999"
    
    def test_multiple_students_with_keys(self, service):
        """Test multiple students collecting keys for the same room."""
        service.collect_key("101", "John Doe")
        service.collect_key("101", "Jane Smith")
        
        # Verify the room was updated
        room = service._get_room("101")
        assert len(room.collected_actions) == 2
        assert room.collected_actions[0].student.name == "John Doe"
        assert room.collected_actions[1].student.name == "Jane Smith"
        
        # Should have no available keys now
        assert room.available_keys == 0
        
        # Attempting to collect another key should fail
        with pytest.raises(NoKeysAvailableException):
            service.collect_key("101", "Bob Brown")
        
        # Return a key
        service.return_key("101", "John Doe")
        
        # Now should have one available key
        assert service._get_room("101").available_keys == 1
        
        # Another student can collect a key
        service.collect_key("101", "Bob Brown")
        
        # Should have no available keys again
        assert service._get_room("101").available_keys == 0
    
    @patch('utils.logger.info')
    def test_logging_collect_key(self, mock_log, service):
        """Test that collecting a key logs the correct messages."""
        service.collect_key("101", "John Doe")
        
        # Verify log calls
        assert mock_log.call_count >= 2
        mock_log.assert_any_call("Student John Doe is collecting a key for room 101")
        mock_log.assert_any_call("Student John Doe collected a key for room 101")
    
    @patch('utils.logger.info')
    def test_logging_return_key(self, mock_log, service):
        """Test that returning a key logs the correct messages."""
        # First collect a key
        service.collect_key("101", "John Doe")
        mock_log.reset_mock()  # Reset the mock to clear previous calls
        
        # Then return it
        service.return_key("101", "John Doe")
        
        # Verify log calls
        assert mock_log.call_count >= 2
        mock_log.assert_any_call("Student John Doe is returning a key for room 101")
        mock_log.assert_any_call("Student John Doe returned a key for room 101")
    
    @patch('utils.logger.info')
    def test_logging_report_lost_key(self, mock_log, service):
        """Test that reporting a lost key logs the correct messages."""
        # Room 103 already has a collected key for John Doe
        mock_log.reset_mock()  # Reset the mock to clear previous calls
        
        service.report_lost_key("103", "John Doe")
        
        # Verify log calls
        assert mock_log.call_count >= 2
        mock_log.assert_any_call("Student John Doe is reporting a lost key for room 103")
        mock_log.assert_any_call("Student John Doe reported a lost key for room 103")
    
    @patch('utils.logger.info')
    def test_logging_borrow_spare_key(self, mock_log, service):
        """Test that borrowing a spare key logs the correct messages."""
        mock_log.reset_mock()  # Reset the mock to clear previous calls
        
        service.borrow_spare_key("101", "Jane Smith")
        
        # Verify log calls
        assert mock_log.call_count >= 2
        mock_log.assert_any_call("Student Jane Smith is borrowing a spare key for room 101")
        mock_log.assert_any_call("Student Jane Smith borrowed a spare key for room 101")