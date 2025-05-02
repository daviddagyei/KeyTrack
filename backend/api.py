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
                # Get counts of different actions
                collected_count = len(room_data.get('collected', []))
                returned_count = len(room_data.get('returned', []))
                lost_count = len(room_data.get('lost', []))
                borrowed_count = len(room_data.get('borrowed', []))
                
                # Calculate active collected keys (collected minus returned)
                # Returns are now accounted against both collected and borrowed keys
                active_collected = max(0, collected_count - returned_count)
                active_borrowed = max(0, borrowed_count - max(0, returned_count - collected_count))
                
                formatted_room = {
                    "id": room_id,
                    "total_keys": room_data.get('total_keys', 0),
                    "available_keys": room_data.get('available_keys', 0),
                    "collected_keys": active_collected,
                    "lost_keys": lost_count,
                    "borrowed_keys": active_borrowed,
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
            # Use the robust room ID normalization
            normalized_room_id = self._normalize_room_id(room_id)
            rooms_data = self.key_service.get_all_rooms_data()
            
            # Try to find the room with the normalized ID
            if normalized_room_id in rooms_data:
                room_data = rooms_data[normalized_room_id]
            else:
                # If still not found, provide debugging info
                available_rooms = list(map(str, list(rooms_data.keys())[:5]))
                print(f"Room not found. Available rooms (first 5): {available_rooms}")
                print(f"Could not find room with normalized ID: '{normalized_room_id}' (original: '{room_id}')")
                return {
                    "success": False,
                    "error": f"Room {room_id} not found"
                }
            
            # Get counts of different actions
            collected_count = len(room_data.get('collected', []))
            returned_count = len(room_data.get('returned', []))
            lost_count = len(room_data.get('lost', []))
            borrowed_count = len(room_data.get('borrowed', []))
            
            # Calculate active collected and borrowed keys with merged return logic
            active_collected = max(0, collected_count - returned_count)
            active_borrowed = max(0, borrowed_count - max(0, returned_count - collected_count))
            
            formatted_room = {
                "id": room_id,  # Use the original room_id for display purposes
                "total_keys": room_data.get('total_keys', 0),
                "available_keys": room_data.get('available_keys', 0),
                "collected_keys": active_collected,
                "lost_keys": lost_count,
                "borrowed_keys": active_borrowed,
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
            print(f"Error in get_room: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def _normalize_room_id(self, room_id):
        """Helper method to normalize room IDs to the format stored in the database"""
        try:
            # First convert room_id to string if it's not already
            room_id = str(room_id)
            
            # Remove any URL encoding (like %20)
            if '%20' in room_id:
                room_id = room_id.replace('%20', ' ')
            
            # Strip any whitespace
            room_id = room_id.strip()
            
            # Get all available rooms
            rooms_data = self.key_service.get_all_rooms_data()
            
            # Try direct case-insensitive match
            for key in rooms_data.keys():
                # Convert key to string to avoid 'int' has no attribute 'lower' error
                str_key = str(key)
                if str_key.lower() == room_id.lower():
                    return key
                
            # Try with "Room " prefix - for when user inputs just the number
            room_with_prefix = f"Room {room_id}"
            for key in rooms_data.keys():
                # Convert key to string
                str_key = str(key)
                if str_key.lower() == room_with_prefix.lower():
                    return key
                
            # Try without "Room " prefix - for when the system expects just the number
            if room_id.lower().startswith("room "):
                room_without_prefix = room_id[5:].strip()
                for key in rooms_data.keys():
                    # Convert key to string
                    str_key = str(key)
                    if str_key.lower() == room_without_prefix.lower():
                        return key
            
            # If still not found, try extracting just the numeric part
            import re
            numeric_match = re.search(r'\d+', room_id)
            if numeric_match:
                number = numeric_match.group()
                # Try just the number
                for key in rooms_data.keys():
                    str_key = str(key)
                    if str_key == number or str_key.endswith(f" {number}"):
                        return key
            
            # Print available rooms for debugging (limit to first 5)
            available_rooms = list(map(str, list(rooms_data.keys())[:5]))
            print(f"Available rooms (first 5): {available_rooms}")
            print(f"Could not find a match for room ID: '{room_id}'")
            
            # No match found, return original and let the service handle the error
            return room_id
        except Exception as e:
            print(f"Error in _normalize_room_id: {e}")
            # If any error occurs, return the original ID
            return room_id
    
    def collect_key(self, room_id, student_name):
        """Collect a key for a specific room"""
        try:
            normalized_room_id = self._normalize_room_id(room_id)
            self.key_service.collect_key(normalized_room_id, student_name)
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
            normalized_room_id = self._normalize_room_id(room_id)
            self.key_service.return_key(normalized_room_id, student_name)
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
            normalized_room_id = self._normalize_room_id(room_id)
            self.key_service.report_lost_key(normalized_room_id, student_name)
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
            normalized_room_id = self._normalize_room_id(room_id)
            self.key_service.borrow_spare_key(normalized_room_id, student_name)
            return {
                "success": True,
                "message": f"Spare key for room {room_id} borrowed by {student_name}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }

    def return_borrowed_key(self, room_id, student_name):
        """Return a borrowed key for a specific room"""
        try:
            normalized_room_id = self._normalize_room_id(room_id)
            self.key_service.return_borrowed_key(normalized_room_id, student_name)
            return {
                "success": True,
                "message": f"Borrowed key for room {room_id} returned by {student_name}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }