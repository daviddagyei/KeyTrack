from datetime import datetime
from typing import List, Optional
from models.base import Entity


class Student(Entity):
    """Represents a student in the key tracking system."""
    
    def __init__(self, id: str, name: str):
        super().__init__(id)
        self.name = name
    
    def __str__(self) -> str:
        return f"Student({self.id}: {self.name})"


class KeyAction:
    """Represents an action performed on a key."""
    
    def __init__(self, student: Student, timestamp: datetime):
        self.student = student
        self.timestamp = timestamp
    
    def __str__(self) -> str:
        return f"{self.student.name} ({self.timestamp.strftime('%Y-%m-%d %H:%M:%S')})"
    
    @classmethod
    def from_string(cls, action_str: str, student_id_generator) -> 'KeyAction':
        """Parse a key action from string format."""
        # Format: "Student Name (YYYY-MM-DD HH:MM:SS)"
        name_part, timestamp_part = action_str.rsplit(" (", 1)
        timestamp = datetime.strptime(timestamp_part.rstrip(")"), "%Y-%m-%d %H:%M:%S")
        student = Student(student_id_generator(name_part), name_part)
        return cls(student, timestamp)


class Room(Entity):
    """Represents a room with keys."""
    
    def __init__(self, id: str, total_keys: int = 0):
        super().__init__(id)
        self.total_keys = total_keys
        self.collected_actions: List[KeyAction] = []
        self.lost_actions: List[KeyAction] = []
        self.borrowed_actions: List[KeyAction] = []
        self.returned_actions: List[KeyAction] = []
    
    @property
    def available_keys(self) -> int:
        """Calculate the number of available keys."""
        collected = len(self.collected_actions)
        returned = len(self.returned_actions)
        lost = len(self.lost_actions)
        # Total keys minus currently out (collected-returned) minus lost
        return self.total_keys - (collected - returned) - lost
    
    def has_key(self, student_name: str) -> bool:
        """Check if a student has a key for this room.
        
        This accounts for students who may have collected multiple keys
        and returned some but not all of them.
        """
        # Count how many keys the student has collected
        collected_count = sum(1 for action in self.collected_actions 
                             if action.student.name == student_name)
        
        # Count how many keys the student has returned
        returned_count = sum(1 for action in self.returned_actions 
                            if action.student.name == student_name)
        
        # If the student has collected more keys than they've returned, they still have a key
        return collected_count > returned_count