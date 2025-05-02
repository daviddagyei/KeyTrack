from flask import Flask, jsonify, request
from services.key_management_service import DefaultKeyManagementService
from repository.excel_repository import ExcelRepository
from backend.api import KeyTrackAPI
import os

def create_app():
    app = Flask(__name__)
    
    # Initialize the key management service with an absolute path to the Excel file
    # This ensures the file is found regardless of the current working directory
    app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    excel_path = os.path.join(app_root, 'key_distribution.xlsx')
    repository = ExcelRepository(excel_path)
    service = DefaultKeyManagementService(repository)
    api = KeyTrackAPI(service)
    
    # Define API routes
    @app.route('/api/rooms', methods=['GET'])
    def get_rooms():
        return jsonify(api.get_rooms())
    
    @app.route('/api/rooms/<room_id>', methods=['GET'])
    def get_room(room_id):
        return jsonify(api.get_room(room_id))
    
    @app.route('/api/actions/collect', methods=['POST'])
    def collect_key():
        data = request.json
        room_id = data.get('room_id')
        student_name = data.get('student_name')
        
        if not room_id or not student_name:
            return jsonify({
                "success": False,
                "error": "Missing room_id or student_name"
            }), 400
        
        return jsonify(api.collect_key(room_id, student_name))
    
    @app.route('/api/actions/return', methods=['POST'])
    def return_key():
        data = request.json
        room_id = data.get('room_id')
        student_name = data.get('student_name')
        
        if not room_id or not student_name:
            return jsonify({
                "success": False,
                "error": "Missing room_id or student_name"
            }), 400
        
        return jsonify(api.return_key(room_id, student_name))
    
    @app.route('/api/actions/lost', methods=['POST'])
    def report_lost_key():
        data = request.json
        room_id = data.get('room_id')
        student_name = data.get('student_name')
        
        if not room_id or not student_name:
            return jsonify({
                "success": False,
                "error": "Missing room_id or student_name"
            }), 400
        
        return jsonify(api.report_lost_key(room_id, student_name))
    
    @app.route('/api/actions/borrow', methods=['POST'])
    def borrow_spare_key():
        data = request.json
        room_id = data.get('room_id')
        student_name = data.get('student_name')
        
        if not room_id or not student_name:
            return jsonify({
                "success": False,
                "error": "Missing room_id or student_name"
            }), 400
        
        return jsonify(api.borrow_spare_key(room_id, student_name))
    
    @app.route('/api/student/<student_name>', methods=['GET'])
    def get_student_history(student_name):
        """Get all actions performed by a specific student"""
        rooms = api.get_rooms()
        if not rooms.get('success'):
            return jsonify({
                "success": False,
                "error": "Failed to retrieve rooms data"
            }), 500
        
        history = []
        for room in rooms['rooms']:
            # Check collected actions
            for action in room.get('collected_actions', []):
                if action['student'].lower() == student_name.lower():
                    history.append({
                        "room_id": room['id'],
                        "action_type": "collected",
                        "timestamp": action['timestamp']
                    })
            
            # Check returned actions
            for action in room.get('returned_actions', []):
                if action['student'].lower() == student_name.lower():
                    history.append({
                        "room_id": room['id'],
                        "action_type": "returned",
                        "timestamp": action['timestamp']
                    })
            
            # Check lost actions
            for action in room.get('lost_actions', []):
                if action['student'].lower() == student_name.lower():
                    history.append({
                        "room_id": room['id'],
                        "action_type": "reported_lost",
                        "timestamp": action['timestamp']
                    })
            
            # Check borrowed actions
            for action in room.get('borrowed_actions', []):
                if action['student'].lower() == student_name.lower():
                    history.append({
                        "room_id": room['id'],
                        "action_type": "borrowed",
                        "timestamp": action['timestamp']
                    })
        
        # Sort history by timestamp (newest first)
        history = sorted(history, key=lambda x: x['timestamp'], reverse=True)
        
        return jsonify({
            "success": True,
            "student_name": student_name,
            "history": history
        })
    
    @app.route('/api/keys/statistics', methods=['GET'])
    def get_key_statistics():
        """Get overall statistics about keys"""
        try:
            # Get all rooms data
            rooms_data = api.get_rooms()
            
            if not rooms_data.get('success'):
                return jsonify({
                    "success": False,
                    "error": "Failed to retrieve rooms data"
                }), 500
                
            rooms = rooms_data.get('rooms', [])
            
            # Calculate statistics
            total_keys = 0
            available_keys = 0
            keys_out = 0
            keys_lost = 0
            keys_borrowed = 0
            
            for room in rooms:
                total_keys += room.get('total_keys', 0)
                available_keys += room.get('available_keys', 0)
                keys_out += room.get('collected_keys', 0)
                keys_lost += room.get('lost_keys', 0)
                keys_borrowed += room.get('borrowed_keys', 0)
            
            return jsonify({
                "success": True,
                "total_keys": total_keys,
                "available_keys": available_keys,
                "keys_out": keys_out,
                "keys_lost": keys_lost,
                "keys_borrowed": keys_borrowed,
                "keys_damaged": 0  # Not tracking damaged keys yet
            })
        except Exception as e:
            return jsonify({
                "success": False,
                "error": str(e)
            }), 500
    
    @app.route('/api/rooms/overview', methods=['GET'])
    def get_rooms_overview():
        """Get overview data for all rooms"""
        try:
            # Get all rooms data
            rooms_data = api.get_rooms()
            
            if not rooms_data.get('success'):
                return jsonify({
                    "success": False,
                    "error": "Failed to retrieve rooms data"
                }), 500
            
            # Format for the dashboard
            rooms = []
            for room in rooms_data.get('rooms', []):
                rooms.append({
                    "room_id": room.get('id', ''),
                    "total_keys": room.get('total_keys', 0),
                    "available_keys": room.get('available_keys', 0),
                    "collected_keys": room.get('collected_keys', 0),
                    "lost_keys": room.get('lost_keys', 0),
                    "borrowed_keys": room.get('borrowed_keys', 0)
                })
            
            return jsonify(rooms)
        except Exception as e:
            return jsonify({
                "success": False,
                "error": str(e)
            }), 500
    
    @app.route('/api/activity/recent', methods=['GET'])
    def get_recent_activity():
        """Get recent key activity"""
        try:
            # Get all rooms data
            rooms_data = api.get_rooms()
            
            if not rooms_data.get('success'):
                return jsonify({
                    "success": False,
                    "error": "Failed to retrieve rooms data"
                }), 500
            
            # Compile all activities
            all_activities = []
            
            for room in rooms_data.get('rooms', []):
                room_id = room.get('id', '')
                
                # Add collected actions
                for action in room.get('collected_actions', []):
                    all_activities.append({
                        "room_id": room_id,
                        "student": action.get('student', 'Unknown'),
                        "action": "collected",
                        "description": f"{action.get('student', 'Unknown')} collected a key for Room {room_id}",
                        "timestamp": action.get('timestamp')
                    })
                
                # Add returned actions
                for action in room.get('returned_actions', []):
                    all_activities.append({
                        "room_id": room_id,
                        "student": action.get('student', 'Unknown'),
                        "action": "returned",
                        "description": f"{action.get('student', 'Unknown')} returned a key for Room {room_id}",
                        "timestamp": action.get('timestamp')
                    })
                
                # Add lost actions
                for action in room.get('lost_actions', []):
                    all_activities.append({
                        "room_id": room_id,
                        "student": action.get('student', 'Unknown'),
                        "action": "lost",
                        "description": f"{action.get('student', 'Unknown')} reported a lost key for Room {room_id}",
                        "timestamp": action.get('timestamp')
                    })
                
                # Add borrowed actions
                for action in room.get('borrowed_actions', []):
                    all_activities.append({
                        "room_id": room_id,
                        "student": action.get('student', 'Unknown'),
                        "action": "borrowed",
                        "description": f"{action.get('student', 'Unknown')} borrowed a spare key for Room {room_id}",
                        "timestamp": action.get('timestamp')
                    })
            
            # Sort activities by timestamp (newest first)
            all_activities = sorted(all_activities, key=lambda x: x.get('timestamp', ''), reverse=True)
            
            # Limit to 10 most recent activities
            recent_activities = all_activities[:10]
            
            return jsonify(recent_activities)
        except Exception as e:
            return jsonify({
                "success": False,
                "error": str(e)
            }), 500
    
    return app