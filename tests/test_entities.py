import pytest
from datetime import datetime
from models.entities import Room, Student, KeyAction


class TestStudent:
    """Test cases for the Student entity."""
    
    def test_student_initialization(self):
        """Test that a Student object initializes with correct attributes."""
        student = Student(id="S_1234", name="John Doe")
        assert student.id == "S_1234"
        assert student.name == "John Doe"
    
    def test_student_id_property(self):
        """Test that the student ID property works correctly."""
        student = Student(id="S_1234", name="John Doe")
        assert student.id == "S_1234"
    
    def test_student_name_property(self):
        """Test that the student name property works correctly."""
        student = Student(id="S_1234", name="John Doe")
        assert student.name == "John Doe"
    
    def test_student_equality_with_same_id_different_name(self):
        """Test that students with the same ID are equal even with different names."""
        student1 = Student(id="S_1234", name="John Doe")
        student2 = Student(id="S_1234", name="Jane Smith")
        assert student1 == student2
    
    def test_student_inequality_with_different_id_same_name(self):
        """Test that students with different IDs are not equal even with the same name."""
        student1 = Student(id="S_1234", name="John Doe")
        student2 = Student(id="S_5678", name="John Doe")
        assert student1 != student2
    
    def test_student_equality_with_same_id_and_name(self):
        """Test that students with the same ID and name are equal."""
        student1 = Student(id="S_1234", name="John Doe")
        student2 = Student(id="S_1234", name="John Doe")
        assert student1 == student2
    
    def test_student_self_equality(self):
        """Test that a student is equal to itself."""
        student = Student(id="S_1234", name="John Doe")
        assert student == student
    
    def test_student_inequality_with_non_student_object(self):
        """Test that a student is not equal to a non-student object."""
        student = Student(id="S_1234", name="John Doe")
        assert student != "S_1234"
        assert student != 123
        assert student != ["S_1234"]
        assert student != {"id": "S_1234", "name": "John Doe"}
    
    def test_student_string_representation(self):
        """Test string representation of a student."""
        student = Student(id="S_1234", name="John Doe")
        assert str(student) == "Student(S_1234: John Doe)"
    
    def test_student_string_representation_with_special_chars(self):
        """Test string representation of a student with special characters in the name."""
        student = Student(id="S_1234", name="John & Mary O'Doe")
        assert str(student) == "Student(S_1234: John & Mary O'Doe)"


class TestKeyAction:
    """Test cases for the KeyAction entity."""
    
    def test_key_action_initialization(self, sample_student):
        """Test creating a KeyAction instance with basic attributes."""
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action = KeyAction(student=sample_student, timestamp=timestamp)
        
        assert action.student == sample_student
        assert action.timestamp == timestamp
    
    def test_key_action_student_property(self, sample_student):
        """Test the student property of KeyAction."""
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action = KeyAction(student=sample_student, timestamp=timestamp)
        
        assert action.student == sample_student
        assert action.student.name == "Test Student"
        assert action.student.id == "S_1234"
    
    def test_key_action_timestamp_property(self, sample_student):
        """Test the timestamp property of KeyAction."""
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action = KeyAction(student=sample_student, timestamp=timestamp)
        
        assert action.timestamp == timestamp
        assert action.timestamp.year == 2025
        assert action.timestamp.month == 5
        assert action.timestamp.day == 1
        assert action.timestamp.hour == 10
        assert action.timestamp.minute == 30
    
    def test_key_action_string_representation(self, sample_student):
        """Test string representation of a key action."""
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action = KeyAction(student=sample_student, timestamp=timestamp)
        
        assert str(action) == "Test Student (2025-05-01 10:30:00)"
    
    def test_key_action_string_representation_different_timezone(self, sample_student):
        """Test string representation with a different timestamp."""
        timestamp = datetime(2025, 12, 31, 23, 59, 59)
        action = KeyAction(student=sample_student, timestamp=timestamp)
        
        assert str(action) == "Test Student (2025-12-31 23:59:59)"
    
    def test_key_action_from_string_simple(self):
        """Test creating a KeyAction from a basic string representation."""
        action_string = "John Doe (2025-05-01 10:30:00)"
        
        # Mock student_id_generator
        id_generator = lambda name: f"S_{hash(name) % 10000}"
        
        action = KeyAction.from_string(action_string, id_generator)
        
        assert action.student.name == "John Doe"
        assert action.timestamp == datetime(2025, 5, 1, 10, 30, 0)
    
    def test_key_action_from_string_with_spaces(self):
        """Test creating a KeyAction from a string with extra spaces in the name."""
        action_string = "John  Doe  Smith (2025-05-01 10:30:00)"
        
        # Mock student_id_generator
        id_generator = lambda name: f"S_{hash(name) % 10000}"
        
        action = KeyAction.from_string(action_string, id_generator)
        
        assert action.student.name == "John  Doe  Smith"
        assert action.timestamp == datetime(2025, 5, 1, 10, 30, 0)
    
    def test_key_action_from_string_with_special_chars(self):
        """Test creating a KeyAction from a string with special characters."""
        action_string = "O'Connor & D'Arcy (2025-05-01 10:30:00)"
        
        # Mock student_id_generator
        id_generator = lambda name: f"S_{hash(name) % 10000}"
        
        action = KeyAction.from_string(action_string, id_generator)
        
        assert action.student.name == "O'Connor & D'Arcy"
        assert action.timestamp == datetime(2025, 5, 1, 10, 30, 0)
    
    def test_key_action_from_string_id_generation(self):
        """Test that student ID is generated correctly from string."""
        action_string = "John Doe (2025-05-01 10:30:00)"
        
        # Use a predictable ID generator for testing
        id_generator = lambda name: f"S_TEST_{name}"
        
        action = KeyAction.from_string(action_string, id_generator)
        
        assert action.student.id == "S_TEST_John Doe"
        assert action.student.name == "John Doe"
    
    def test_key_action_from_string_with_parentheses_in_name(self):
        """Test creating a KeyAction from a string with parentheses in the name."""
        action_string = "John (Nickname) Doe (2025-05-01 10:30:00)"
        
        # Mock student_id_generator
        id_generator = lambda name: f"S_{hash(name) % 10000}"
        
        action = KeyAction.from_string(action_string, id_generator)
        
        assert action.student.name == "John (Nickname) Doe"
        assert action.timestamp == datetime(2025, 5, 1, 10, 30, 0)


class TestRoom:
    """Test cases for the Room entity."""
    
    def test_room_initialization_basics(self):
        """Test creating a Room instance with basic attributes."""
        room = Room(id="101", total_keys=5)
        
        assert room.id == "101"
        assert room.total_keys == 5
        assert room.available_keys == 5
    
    def test_room_initialization_collections(self):
        """Test that the Room's collections initialize as empty lists."""
        room = Room(id="101", total_keys=5)
        
        assert isinstance(room.collected_actions, list)
        assert len(room.collected_actions) == 0
        
        assert isinstance(room.lost_actions, list)
        assert len(room.lost_actions) == 0
        
        assert isinstance(room.borrowed_actions, list)
        assert len(room.borrowed_actions) == 0
        
        assert isinstance(room.returned_actions, list)
        assert len(room.returned_actions) == 0
    
    def test_room_id_property(self):
        """Test the id property of Room."""
        room = Room(id="101", total_keys=5)
        assert room.id == "101"
    
    def test_room_total_keys_property(self):
        """Test the total_keys property of Room."""
        room = Room(id="101", total_keys=5)
        assert room.total_keys == 5
    
    def test_room_available_keys_with_no_actions(self):
        """Test the available_keys property calculation with no actions."""
        room = Room(id="101", total_keys=5)
        assert room.available_keys == 5
    
    def test_room_available_keys_with_collected_keys(self, sample_student):
        """Test available_keys calculation with collected keys."""
        room = Room(id="101", total_keys=5)
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action = KeyAction(student=sample_student, timestamp=timestamp)
        
        room.collected_actions.append(action)
        assert room.available_keys == 4
        
        room.collected_actions.append(action)
        assert room.available_keys == 3
    
    def test_room_available_keys_with_returned_keys(self, sample_student):
        """Test available_keys calculation with returned keys."""
        room = Room(id="101", total_keys=5)
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action = KeyAction(student=sample_student, timestamp=timestamp)
        
        room.collected_actions.append(action)
        assert room.available_keys == 4
        
        room.returned_actions.append(action)
        assert room.available_keys == 5
    
    def test_room_available_keys_with_lost_keys(self, sample_student):
        """Test available_keys calculation with lost keys."""
        room = Room(id="101", total_keys=5)
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action = KeyAction(student=sample_student, timestamp=timestamp)
        
        room.lost_actions.append(action)
        assert room.available_keys == 4
    
    def test_room_available_keys_with_collected_and_lost_keys(self, sample_student):
        """Test available_keys with both collected and lost keys."""
        room = Room(id="101", total_keys=5)
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action1 = KeyAction(student=sample_student, timestamp=timestamp)
        action2 = KeyAction(student=Student(id="S_5678", name="Jane Smith"), timestamp=timestamp)
        
        room.collected_actions.append(action1)
        room.lost_actions.append(action2)
        
        assert room.available_keys == 3
    
    def test_room_available_keys_complex_scenario(self, sample_student):
        """Test available_keys with a complex scenario of actions."""
        room = Room(id="101", total_keys=10)
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        
        # Create students
        student1 = sample_student
        student2 = Student(id="S_5678", name="Jane Smith")
        student3 = Student(id="S_9012", name="Bob Brown")
        
        # Create actions
        action1 = KeyAction(student=student1, timestamp=timestamp)
        action2 = KeyAction(student=student2, timestamp=timestamp)
        action3 = KeyAction(student=student3, timestamp=timestamp)
        
        # Student 1 collects a key
        room.collected_actions.append(action1)
        assert room.available_keys == 9
        
        # Student 2 collects a key
        room.collected_actions.append(action2)
        assert room.available_keys == 8
        
        # Student 3 collects a key
        room.collected_actions.append(action3)
        assert room.available_keys == 7
        
        # Student 1 returns the key
        room.returned_actions.append(action1)
        assert room.available_keys == 8
        
        # Student 2 reports key as lost
        room.lost_actions.append(action2)
        assert room.available_keys == 7
        
        # Student 3 still has the key out
        assert room.available_keys == 7
    
    def test_room_has_key_when_student_doesnt_have_key(self, sample_student):
        """Test has_key returns False when student doesn't have a key."""
        room = Room(id="101", total_keys=5)
        assert room.has_key(sample_student.name) is False
    
    def test_room_has_key_when_student_has_collected_key(self, sample_student):
        """Test has_key returns True when student has collected a key."""
        room = Room(id="101", total_keys=5)
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action = KeyAction(student=sample_student, timestamp=timestamp)
        
        room.collected_actions.append(action)
        assert room.has_key(sample_student.name) is True
    
    def test_room_has_key_when_student_has_returned_key(self, sample_student):
        """Test has_key returns False when student has returned a key."""
        room = Room(id="101", total_keys=5)
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        action = KeyAction(student=sample_student, timestamp=timestamp)
        
        room.collected_actions.append(action)
        room.returned_actions.append(action)
        assert room.has_key(sample_student.name) is False
    
    def test_room_has_key_with_multiple_students(self, sample_student):
        """Test has_key with multiple students having keys."""
        room = Room(id="101", total_keys=5)
        timestamp = datetime(2025, 5, 1, 10, 30, 0)
        
        student1 = sample_student
        student2 = Student(id="S_5678", name="Jane Smith")
        
        action1 = KeyAction(student=student1, timestamp=timestamp)
        action2 = KeyAction(student=student2, timestamp=timestamp)
        
        room.collected_actions.append(action1)
        room.collected_actions.append(action2)
        
        assert room.has_key(student1.name) is True
        assert room.has_key(student2.name) is True
        assert room.has_key("Unknown Student") is False
    
    def test_room_has_key_with_multiple_collections_by_same_student(self, sample_student):
        """Test has_key when a student has collected multiple keys."""
        room = Room(id="101", total_keys=5)
        timestamp1 = datetime(2025, 5, 1, 10, 30, 0)
        timestamp2 = datetime(2025, 5, 1, 11, 30, 0)
        
        action1 = KeyAction(student=sample_student, timestamp=timestamp1)
        action2 = KeyAction(student=sample_student, timestamp=timestamp2)
        
        room.collected_actions.append(action1)
        room.collected_actions.append(action2)
        room.returned_actions.append(action1)  # Return one key
        
        # Student still has one key out
        assert room.has_key(sample_student.name) is True