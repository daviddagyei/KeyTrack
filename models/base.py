from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional
from datetime import datetime


class Entity:
    """Base class for all entities in the system."""
    def __init__(self, id: str):
        self.id = id
    
    def __eq__(self, other):
        if not isinstance(other, Entity):
            return False
        return self.id == other.id


class DataRepository(ABC):
    """Abstract base class for data repositories."""
    
    @abstractmethod
    def load(self) -> Dict[str, Any]:
        """Load data from the storage medium."""
        pass
    
    @abstractmethod
    def save(self, data: Dict[str, Any]) -> bool:
        """Save data to the storage medium."""
        pass


class KeyManagementService(ABC):
    """Abstract base class for key management operations."""
    
    @abstractmethod
    def collect_key(self, room_id: str, student_name: str) -> bool:
        """Handle key collection for a room by a student."""
        pass
    
    @abstractmethod
    def return_key(self, room_id: str, student_name: str) -> bool:
        """Handle key return for a room by a student."""
        pass
    
    @abstractmethod
    def report_lost_key(self, room_id: str, student_name: str) -> bool:
        """Handle reporting of a lost key for a room by a student."""
        pass
    
    @abstractmethod
    def borrow_spare_key(self, room_id: str, student_name: str) -> bool:
        """Handle borrowing of a spare key for a room by a student."""
        pass