from services.key_management_service import KeyManagementService
from datetime import datetime
import json

class KeyTrackAPI:
    """API for the KeyTrack application, provides an interface between the web service and business logic"""
    
    def __init__(self, key_service: KeyManagementService):
        self.key_service = key_service
    
    def get_rooms(self):
        """Get information about all rooms"""
        try:
            rooms_data = self.key_service.get_all_rooms_data()
            
            # Format the response to include various statistics
            rooms = []
            for room_id, room_data in rooms_data.items():
                formatted_room = {
                    "id": room_id,
                    "total_keys": room_data.get('total_keys', 0),
                    "available_keys": room_data.get('available_keys', 0),
                    "collected_keys": len(room_data.get('collected', [])),
                    "lost_keys": len(room_data.get('lost', [])),
                    "borrowed_keys": len(room_data.get('borrowed', [])),
                    "collected_actions": [],
                    "returned_actions": [],
                    "lost_actions": [],
                    "borrowed_actions": []
                }
                
                # Add detailed action records
                if 'history' in room_data:
                    for action in room_data['history']:
                        action_type = action.get('action')
                        if action_type:
                            action_entry = {
                                "student": action.get('student', 'Unknown'),
                                "timestamp": action.get('timestamp', datetime.now().isoformat())
                            }
                            
                            if action_type == 'collected':
                                formatted_room['collected_actions'].append(action_entry)
                            elif action_type == 'returned':
                                formatted_room['returned_actions'].append(action_entry)
                            elif action_type == 'lost':
                                formatted_room['lost_actions'].append(action_entry)
                            elif action_type == 'borrowed':
                                formatted_room['borrowed_actions'].append(action_entry)
                
                rooms.append(formatted_room)
            
            return {
                "success": True,
                "rooms": rooms
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def get_room(self, room_id):
        """Get detailed information about a specific room"""
        try:
            rooms_data = self.key_service.get_all_rooms_data()
            if room_id not in rooms_data:
                return {
                    "success": False,
                    "error": f"Room {room_id} not found"
                }
            
            room_data = rooms_data[room_id]
            formatted_room = {
                "id": room_id,
                "total_keys": room_data.get('total_keys', 0),
                "available_keys": room_data.get('available_keys', 0),
                "collected_keys": len(room_data.get('collected', [])),
                "lost_keys": len(room_data.get('lost', [])),
                "borrowed_keys": len(room_data.get('borrowed', [])),
                "collected_actions": [],
                "returned_actions": [],
                "lost_actions": [],
                "borrowed_actions": []
            }
            
            # Add detailed action records
            if 'history' in room_data:
                for action in room_data['history']:
                    action_type = action.get('action')
                    if action_type:
                        action_entry = {
                            "student": action.get('student', 'Unknown'),
                            "timestamp": action.get('timestamp', datetime.now().isoformat())
                        }
                        
                        if action_type == 'collected':
                            formatted_room['collected_actions'].append(action_entry)
                        elif action_type == 'returned':
                            formatted_room['returned_actions'].append(action_entry)
                        elif action_type == 'lost':
                            formatted_room['lost_actions'].append(action_entry)
                        elif action_type == 'borrowed':
                            formatted_room['borrowed_actions'].append(action_entry)
            
            return {
                "success": True,
                "room": formatted_room
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def collect_key(self, room_id, student_name):
        """Collect a key for a specific room"""
        try:
            self.key_service.collect_key(room_id, student_name)
            return {
                "success": True,
                "message": f"Key for room {room_id} collected by {student_name}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def return_key(self, room_id, student_name):
        """Return a key for a specific room"""
        try:
            self.key_service.return_key(room_id, student_name)
            return {
                "success": True,
                "message": f"Key for room {room_id} returned by {student_name}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def report_lost_key(self, room_id, student_name):
        """Report a lost key for a specific room"""
        try:
            self.key_service.report_lost_key(room_id, student_name)
            return {
                "success": True,
                "message": f"Key for room {room_id} reported lost by {student_name}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def borrow_spare_key(self, room_id, student_name):
        """Borrow a spare key for a specific room"""
        try:
            self.key_service.borrow_spare_key(room_id, student_name)
            return {
                "success": True,
                "message": f"Spare key for room {room_id} borrowed by {student_name}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }