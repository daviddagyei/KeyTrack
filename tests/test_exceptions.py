import pytest
from models.exceptions import (
    KeyTrackException,
    RoomNotFoundException,
    NoKeysAvailableException,
    KeyNotCollectedException,
    DataAccessException
)


class TestKeyTrackException:
    """Test cases for the base KeyTrackException."""
    
    def test_keytrack_exception_instantiation(self):
        """Test instantiating a KeyTrackException."""
        ex = KeyTrackException("Test message")
        assert isinstance(ex, Exception)
        assert str(ex) == "Test message"
    
    def test_keytrack_exception_inheritance(self):
        """Test that KeyTrackException inherits from Exception."""
        ex = KeyTrackException("Test message")
        assert isinstance(ex, Exception)
        assert issubclass(KeyTrackException, Exception)
    
    def test_keytrack_exception_raising(self):
        """Test raising a KeyTrackException."""
        with pytest.raises(KeyTrackException) as exc_info:
            raise KeyTrackException("Test raising")
        assert str(exc_info.value) == "Test raising"
    
    def test_keytrack_exception_catching_as_parent(self):
        """Test that KeyTrackException can be caught as Exception."""
        try:
            raise KeyTrackException("Test catching")
            assert False, "Exception was not raised"
        except Exception as e:
            assert str(e) == "Test catching"
    
    def test_keytrack_exception_with_empty_message(self):
        """Test KeyTrackException with an empty message."""
        ex = KeyTrackException("")
        assert str(ex) == ""


class TestRoomNotFoundException:
    """Test cases for RoomNotFoundException."""
    
    def test_room_not_found_exception_instantiation(self):
        """Test instantiating a RoomNotFoundException."""
        ex = RoomNotFoundException("101")
        assert isinstance(ex, KeyTrackException)
        assert ex.room_id == "101"
        assert str(ex) == "Room 101 not found."
    
    def test_room_not_found_exception_inheritance(self):
        """Test that RoomNotFoundException inherits from KeyTrackException."""
        ex = RoomNotFoundException("101")
        assert isinstance(ex, KeyTrackException)
        assert issubclass(RoomNotFoundException, KeyTrackException)
    
    def test_room_not_found_exception_raising(self):
        """Test raising a RoomNotFoundException."""
        with pytest.raises(RoomNotFoundException) as exc_info:
            raise RoomNotFoundException("101")
        assert str(exc_info.value) == "Room 101 not found."
        assert exc_info.value.room_id == "101"
    
    def test_room_not_found_exception_catching_as_parent(self):
        """Test that RoomNotFoundException can be caught as KeyTrackException."""
        try:
            raise RoomNotFoundException("101")
            assert False, "Exception was not raised"
        except KeyTrackException as e:
            assert isinstance(e, RoomNotFoundException)
            assert str(e) == "Room 101 not found."
    
    def test_room_not_found_exception_with_special_room_id(self):
        """Test RoomNotFoundException with special characters in room_id."""
        ex = RoomNotFoundException("A-101")
        assert ex.room_id == "A-101"
        assert str(ex) == "Room A-101 not found."
    
    def test_room_not_found_exception_with_numeric_room_id(self):
        """Test RoomNotFoundException with numeric room_id."""
        ex = RoomNotFoundException(101)  # Note: numeric ID
        assert ex.room_id == 101
        assert str(ex) == "Room 101 not found."


class TestNoKeysAvailableException:
    """Test cases for NoKeysAvailableException."""
    
    def test_no_keys_available_exception_instantiation(self):
        """Test instantiating a NoKeysAvailableException."""
        ex = NoKeysAvailableException("101")
        assert isinstance(ex, KeyTrackException)
        assert ex.room_id == "101"
        assert str(ex) == "No keys available for room 101."
    
    def test_no_keys_available_exception_inheritance(self):
        """Test that NoKeysAvailableException inherits from KeyTrackException."""
        ex = NoKeysAvailableException("101")
        assert isinstance(ex, KeyTrackException)
        assert issubclass(NoKeysAvailableException, KeyTrackException)
    
    def test_no_keys_available_exception_raising(self):
        """Test raising a NoKeysAvailableException."""
        with pytest.raises(NoKeysAvailableException) as exc_info:
            raise NoKeysAvailableException("101")
        assert str(exc_info.value) == "No keys available for room 101."
        assert exc_info.value.room_id == "101"
    
    def test_no_keys_available_exception_catching_as_parent(self):
        """Test that NoKeysAvailableException can be caught as KeyTrackException."""
        try:
            raise NoKeysAvailableException("101")
            assert False, "Exception was not raised"
        except KeyTrackException as e:
            assert isinstance(e, NoKeysAvailableException)
            assert str(e) == "No keys available for room 101."
    
    def test_no_keys_available_exception_with_special_room_id(self):
        """Test NoKeysAvailableException with special characters in room_id."""
        ex = NoKeysAvailableException("A-101")
        assert ex.room_id == "A-101"
        assert str(ex) == "No keys available for room A-101."
    
    def test_no_keys_available_exception_with_numeric_room_id(self):
        """Test NoKeysAvailableException with numeric room_id."""
        ex = NoKeysAvailableException(101)  # Note: numeric ID
        assert ex.room_id == 101
        assert str(ex) == "No keys available for room 101."


class TestKeyNotCollectedException:
    """Test cases for KeyNotCollectedException."""
    
    def test_key_not_collected_exception_instantiation(self):
        """Test instantiating a KeyNotCollectedException."""
        ex = KeyNotCollectedException("101", "John Doe")
        assert isinstance(ex, KeyTrackException)
        assert ex.room_id == "101"
        assert ex.student_name == "John Doe"
        assert str(ex) == "No record of key collection for room 101 by student John Doe."
    
    def test_key_not_collected_exception_inheritance(self):
        """Test that KeyNotCollectedException inherits from KeyTrackException."""
        ex = KeyNotCollectedException("101", "John Doe")
        assert isinstance(ex, KeyTrackException)
        assert issubclass(KeyNotCollectedException, KeyTrackException)
    
    def test_key_not_collected_exception_raising(self):
        """Test raising a KeyNotCollectedException."""
        with pytest.raises(KeyNotCollectedException) as exc_info:
            raise KeyNotCollectedException("101", "John Doe")
        assert str(exc_info.value) == "No record of key collection for room 101 by student John Doe."
        assert exc_info.value.room_id == "101"
        assert exc_info.value.student_name == "John Doe"
    
    def test_key_not_collected_exception_catching_as_parent(self):
        """Test that KeyNotCollectedException can be caught as KeyTrackException."""
        try:
            raise KeyNotCollectedException("101", "John Doe")
            assert False, "Exception was not raised"
        except KeyTrackException as e:
            assert isinstance(e, KeyNotCollectedException)
            assert str(e) == "No record of key collection for room 101 by student John Doe."
    
    def test_key_not_collected_exception_with_special_characters(self):
        """Test KeyNotCollectedException with special characters in room_id and student_name."""
        ex = KeyNotCollectedException("A-101", "John O'Connor")
        assert ex.room_id == "A-101"
        assert ex.student_name == "John O'Connor"
        assert str(ex) == "No record of key collection for room A-101 by student John O'Connor."
    
    def test_key_not_collected_exception_with_numeric_room_id(self):
        """Test KeyNotCollectedException with numeric room_id."""
        ex = KeyNotCollectedException(101, "John Doe")  # Note: numeric ID
        assert ex.room_id == 101
        assert ex.student_name == "John Doe"
        assert str(ex) == "No record of key collection for room 101 by student John Doe."


class TestDataAccessException:
    """Test cases for DataAccessException."""
    
    def test_data_access_exception_instantiation(self):
        """Test instantiating a DataAccessException."""
        ex = DataAccessException("Failed to open file")
        assert isinstance(ex, KeyTrackException)
        assert str(ex) == "Failed to open file"
    
    def test_data_access_exception_inheritance(self):
        """Test that DataAccessException inherits from KeyTrackException."""
        ex = DataAccessException("Failed to open file")
        assert isinstance(ex, KeyTrackException)
        assert issubclass(DataAccessException, KeyTrackException)
    
    def test_data_access_exception_raising(self):
        """Test raising a DataAccessException."""
        with pytest.raises(DataAccessException) as exc_info:
            raise DataAccessException("Failed to open file")
        assert str(exc_info.value) == "Failed to open file"
    
    def test_data_access_exception_catching_as_parent(self):
        """Test that DataAccessException can be caught as KeyTrackException."""
        try:
            raise DataAccessException("Failed to open file")
            assert False, "Exception was not raised"
        except KeyTrackException as e:
            assert isinstance(e, DataAccessException)
            assert str(e) == "Failed to open file"
    
    def test_data_access_exception_with_nested_exception(self):
        """Test DataAccessException with a nested exception message."""
        nested_ex = IOError("File not found")
        ex = DataAccessException(f"Database error: {str(nested_ex)}")
        assert str(ex) == "Database error: File not found"
    
    def test_data_access_exception_with_empty_message(self):
        """Test DataAccessException with an empty message."""
        ex = DataAccessException("")
        assert str(ex) == ""