from datetime import datetime
from typing import Dict, Optional, List, Any

from models.base import KeyManagementService, DataRepository
from models.entities import Room, Student, KeyAction
from models.exceptions import (
    RoomNotFoundException,
    NoKeysAvailableException,
    KeyNotCollectedException
)
from utils import logger


class DefaultKeyManagementService(KeyManagementService):
    """Default implementation of the key management service."""
    
    def __init__(self, repository: DataRepository):
        """Initialize the service with a data repository."""
        self.repository = repository
        self._rooms: Optional[Dict[str, Room]] = None
    
    def _load_data(self) -> Dict[str, Room]:
        """Load rooms data from the repository."""
        if self._rooms is None:
            self._rooms = self.repository.load()
        return self._rooms
    
    def _save_data(self) -> bool:
        """Save rooms data to the repository."""
        if self._rooms is not None:
            return self.repository.save(self._rooms)
        return False
    
    def _get_room(self, room_id: str) -> Room:
        """Get a room by ID or raise an exception."""
        rooms = self._load_data()
        if room_id not in rooms:
            raise RoomNotFoundException(room_id)
        return rooms[room_id]
    
    def _create_key_action(self, student_name: str) -> KeyAction:
        """Create a key action for the given student."""
        # Create a simple ID for the student
        student_id = f"S_{hash(student_name) % 10000}"
        student = Student(id=student_id, name=student_name)
        return KeyAction(student=student, timestamp=datetime.now())
    
    def collect_key(self, room_id: str, student_name: str) -> bool:
        """Handle key collection for a room by a student."""
        logger.info(f"Student {student_name} is collecting a key for room {room_id}")
        
        room = self._get_room(room_id)
        
        if room.available_keys <= 0:
            logger.warning(f"No keys available for room {room_id}")
            raise NoKeysAvailableException(room_id)
        
        # Create and add the key action
        action = self._create_key_action(student_name)
        room.collected_actions.append(action)
        
        logger.info(f"Student {student_name} collected a key for room {room_id}")
        return self._save_data()
    
    def return_key(self, room_id: str, student_name: str) -> bool:
        """Handle key return for a room by a student."""
        logger.info(f"Student {student_name} is returning a key for room {room_id}")
        
        room = self._get_room(room_id)
        
        # Check if the student has collected a key
        if not room.has_key(student_name):
            logger.warning(f"No record of key collection for student {student_name} in room {room_id}")
            raise KeyNotCollectedException(room_id, student_name)
        
        # Find the collected key action
        for i, action in enumerate(room.collected_actions):
            if action.student.name == student_name:
                # Create a return action
                return_action = self._create_key_action(student_name)
                room.returned_actions.append(return_action)
                
                logger.info(f"Student {student_name} returned a key for room {room_id}")
                return self._save_data()
        
        # This should not happen if has_key is working correctly
        logger.error(f"Failed to find key collection record for student {student_name} in room {room_id}")
        raise KeyNotCollectedException(room_id, student_name)
    
    def report_lost_key(self, room_id: str, student_name: str) -> bool:
        """Handle reporting of a lost key for a room by a student."""
        logger.info(f"Student {student_name} is reporting a lost key for room {room_id}")
        
        room = self._get_room(room_id)
        
        # Check if the student has collected a key
        if not room.has_key(student_name):
            logger.warning(f"No record of key collection for student {student_name} in room {room_id}")
            raise KeyNotCollectedException(room_id, student_name)
        
        # Find the collected key action
        for i, action in enumerate(room.collected_actions):
            if action.student.name == student_name:
                # Create a lost key action
                lost_action = self._create_key_action(student_name)
                room.lost_actions.append(lost_action)
                
                logger.info(f"Student {student_name} reported a lost key for room {room_id}")
                return self._save_data()
        
        # This should not happen if has_key is working correctly
        logger.error(f"Failed to find key collection record for student {student_name} in room {room_id}")
        raise KeyNotCollectedException(room_id, student_name)
    
    def borrow_spare_key(self, room_id: str, student_name: str) -> bool:
        """Handle borrowing of a spare key for a room by a student."""
        logger.info(f"Student {student_name} is borrowing a spare key for room {room_id}")
        
        room = self._get_room(room_id)
        
        if room.available_keys <= 0:
            logger.warning(f"No spare keys available for room {room_id}")
            raise NoKeysAvailableException(room_id)
        
        # Create and add the key action
        action = self._create_key_action(student_name)
        room.borrowed_actions.append(action)
        
        logger.info(f"Student {student_name} borrowed a spare key for room {room_id}")
        return self._save_data()
    
    def get_all_rooms_data(self) -> Dict[str, Dict[str, Any]]:
        """Get formatted data for all rooms.
        
        Returns:
            Dict[str, Dict[str, Any]]: A dictionary mapping room IDs to room data
        """
        rooms_dict = self._load_data()
        result = {}
        
        for room_id, room in rooms_dict.items():
            room_data = {
                'id': room_id,
                'total_keys': room.total_keys,
                'available_keys': room.available_keys,
                'collected': [],
                'returned': [],
                'lost': [],
                'borrowed': [],
                'history': []
            }
            
            # Add collected actions
            for action in room.collected_actions:
                action_data = {
                    'student': action.student.name,
                    'timestamp': action.timestamp.isoformat(),
                    'action': 'collected'
                }
                room_data['collected'].append(action_data)
                room_data['history'].append(action_data)
            
            # Add returned actions
            for action in room.returned_actions:
                action_data = {
                    'student': action.student.name,
                    'timestamp': action.timestamp.isoformat(),
                    'action': 'returned'
                }
                room_data['returned'].append(action_data)
                room_data['history'].append(action_data)
            
            # Add lost actions
            for action in room.lost_actions:
                action_data = {
                    'student': action.student.name,
                    'timestamp': action.timestamp.isoformat(),
                    'action': 'lost'
                }
                room_data['lost'].append(action_data)
                room_data['history'].append(action_data)
            
            # Add borrowed actions
            for action in room.borrowed_actions:
                action_data = {
                    'student': action.student.name,
                    'timestamp': action.timestamp.isoformat(),
                    'action': 'borrowed'
                }
                room_data['borrowed'].append(action_data)
                room_data['history'].append(action_data)
            
            # Sort history by timestamp (newest first)
            room_data['history'] = sorted(
                room_data['history'],
                key=lambda x: x['timestamp'],
                reverse=True
            )
            
            result[room_id] = room_data
        
        return result