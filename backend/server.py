from flask import Flask, jsonify, request, send_file
from services.key_management_service import DefaultKeyManagementService
from repository.excel_repository import ExcelRepository
from repository.file_manager import ExcelFileManager
from backend.api import KeyTrackAPI
import os
from werkzeug.utils import secure_filename
import tempfile
import shutil
import uuid
from datetime import datetime

# Create a global holder for our service components
# This allows us to update these objects at runtime when the Excel file changes
class ServiceComponents:
    def __init__(self):
        self.app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.excel_path = os.path.join(self.app_root, 'key_distribution.xlsx')
        self.repository = ExcelRepository(self.excel_path)
        self.service = DefaultKeyManagementService(self.repository)
        self.api = KeyTrackAPI(self.service)
        self.file_manager = ExcelFileManager(self.app_root)
        
        # Initialize with the default Excel file if needed
        self.file_manager.initialize_with_default()
        
    def reload(self):
        """Reload the repository with the current Excel file"""
        self.repository = ExcelRepository(self.excel_path)
        self.service = DefaultKeyManagementService(self.repository)
        self.api = KeyTrackAPI(self.service)

# Initialize the global service components
service_components = ServiceComponents()

def create_app():
    app = Flask(__name__)
    
    # Define a temp directory for uploaded files
    app_root = service_components.app_root
    upload_dir = os.path.join(app_root, 'temp_uploads')
    os.makedirs(upload_dir, exist_ok=True)
    
    # Define API routes
    @app.route('/api/rooms', methods=['GET'])
    def get_rooms():
        return jsonify(service_components.api.get_rooms())
    
    @app.route('/api/rooms/<room_id>', methods=['GET'])
    def get_room(room_id):
        return jsonify(service_components.api.get_room(room_id))
    
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
        
        return jsonify(service_components.api.collect_key(room_id, student_name))
    
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
        
        return jsonify(service_components.api.return_key(room_id, student_name))
    
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
        
        return jsonify(service_components.api.report_lost_key(room_id, student_name))
    
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
        
        return jsonify(service_components.api.borrow_spare_key(room_id, student_name))
    
    @app.route('/api/actions/return-borrowed', methods=['POST'])
    def return_borrowed_key():
        data = request.json
        room_id = data.get('room_id')
        student_name = data.get('student_name')
        
        if not room_id or not student_name:
            return jsonify({
                "success": False,
                "error": "Missing room_id or student_name"
            }), 400
        
        return jsonify(service_components.api.return_borrowed_key(room_id, student_name))
    
    @app.route('/api/student/<student_name>', methods=['GET'])
    def get_student_history(student_name):
        """Get all actions performed by a specific student"""
        rooms = service_components.api.get_rooms()
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
            rooms_data = service_components.api.get_rooms()
            
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
                # Now using the correctly calculated collected_keys
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
            rooms_data = service_components.api.get_rooms()
            
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
            rooms_data = service_components.api.get_rooms()
            
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

    @app.route('/api/students/suggestions', methods=['GET'])
    def get_student_suggestions():
        """Get student name suggestions based on partial input"""
        try:
            query = request.args.get('q', '').lower()
            if not query or len(query) < 1:
                return jsonify([])
                
            # Get all rooms data
            rooms_data = service_components.api.get_rooms()
            
            if not rooms_data.get('success'):
                return jsonify([])
                
            # Collect all unique student names
            student_names = set()
            
            for room in rooms_data.get('rooms', []):
                # Check all types of actions (collected, returned, lost, borrowed)
                for action_type in ['collected_actions', 'returned_actions', 'lost_actions', 'borrowed_actions']:
                    for action in room.get(action_type, []):
                        student_names.add(action.get('student', ''))
            
            # Filter students by query
            matches = [name for name in student_names if query in name.lower() and name]
            
            # Sort and limit results
            sorted_matches = sorted(matches)[:10]
            
            return jsonify(sorted_matches)
        except Exception as e:
            print(f"Error in get_student_suggestions: {e}")
            return jsonify([])

    @app.route('/api/excel/download', methods=['GET'])
    def download_excel():
        """Download the current Excel file"""
        try:
            excel_path = service_components.excel_path
            
            # Ensure the file exists
            if not os.path.exists(excel_path):
                return jsonify({
                    "success": False,
                    "error": "Excel file not found"
                }), 404
                
            return send_file(
                excel_path,
                as_attachment=True,
                download_name='key_distribution.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            return jsonify({
                "success": False,
                "error": str(e)
            }), 500
    
    @app.route('/api/excel/upload', methods=['POST'])
    def upload_excel():
        """Upload a new Excel file"""
        try:
            # Check if we should process the file immediately or wait for confirmation
            process_immediately = request.args.get('process', 'false').lower() == 'true'
            
            # Check if the post request has the file part
            if 'file' not in request.files:
                return jsonify({
                    "success": False,
                    "error": "No file part in the request"
                }), 400
                
            file = request.files['file']
            
            # Get optional title
            title = request.form.get('title', '')
            
            # If user does not select file, browser might submit an empty part without filename
            if file.filename == '':
                return jsonify({
                    "success": False,
                    "error": "No file selected"
                }), 400
            
            if file and file.filename.endswith('.xlsx'):
                # Add the file to our file manager
                file_metadata = service_components.file_manager.add_file(file, title)
                
                # Process immediately if requested
                if process_immediately:
                    service_components.file_manager.set_active_file(file_metadata['id'])
                    service_components.reload()
                
                return jsonify({
                    "success": True,
                    "message": "Excel file uploaded successfully",
                    "file": file_metadata
                })
            else:
                return jsonify({
                    "success": False,
                    "error": "Only .xlsx files are allowed"
                }), 400
                
        except Exception as e:
            return jsonify({
                "success": False,
                "error": str(e)
            }), 500
    
    @app.route('/api/excel/files', methods=['GET'])
    def get_excel_files():
        """Get a list of all tracked Excel files"""
        try:
            metadata = service_components.file_manager.get_metadata()
            
            # Add formatted timestamps for display
            for file in metadata['files']:
                try:
                    timestamp = datetime.fromisoformat(file['upload_date'])
                    file['formatted_date'] = timestamp.strftime('%Y-%m-%d %H:%M:%S')
                except:
                    file['formatted_date'] = file['upload_date']
            
            return jsonify({
                "success": True,
                "files": metadata['files'],
                "active_file": metadata['active_file']
            })
        except Exception as e:
            return jsonify({
                "success": False,
                "error": str(e)
            }), 500
    
    @app.route('/api/excel/activate', methods=['POST'])
    def activate_excel():
        """Set an Excel file as the active file"""
        try:
            data = request.json
            file_id = data.get('file_id')
            
            if not file_id:
                return jsonify({
                    "success": False,
                    "error": "No file ID provided"
                }), 400
                
            # Set as active and reload
            result = service_components.file_manager.set_active_file(file_id)
            if result:
                service_components.reload()
                
                return jsonify({
                    "success": True,
                    "message": "Excel file activated successfully"
                })
            else:
                return jsonify({
                    "success": False,
                    "error": "File not found"
                }), 404
            
        except Exception as e:
            return jsonify({
                "success": False,
                "error": str(e)
            }), 500
            
    @app.route('/api/excel/remove', methods=['POST'])
    def remove_excel():
        """Remove an Excel file from tracking"""
        try:
            data = request.json
            file_id = data.get('file_id')
            
            if not file_id:
                return jsonify({
                    "success": False,
                    "error": "No file ID provided"
                }), 400
                
            # Remove the file
            result = service_components.file_manager.remove_file(file_id)
            
            if result:
                return jsonify({
                    "success": True,
                    "message": "Excel file removed successfully"
                })
            else:
                return jsonify({
                    "success": False,
                    "error": "File not found"
                }), 404
            
        except Exception as e:
            return jsonify({
                "success": False,
                "error": str(e)
            }), 500
    
    @app.route('/api/excel/download/<file_id>', methods=['GET'])
    def download_specific_excel(file_id):
        """Download a specific Excel file by ID"""
        try:
            file_path = service_components.file_manager.get_file_path(file_id)
            
            if not file_path or not os.path.exists(file_path):
                return jsonify({
                    "success": False,
                    "error": "File not found"
                }), 404
                
            # Get metadata to use the original filename
            metadata = service_components.file_manager.get_metadata()
            original_name = "excel_file.xlsx"
            
            for file in metadata['files']:
                if file['id'] == file_id:
                    original_name = file['original_name']
                    break
                
            return send_file(
                file_path,
                as_attachment=True,
                download_name=original_name,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            return jsonify({
                "success": False,
                "error": str(e)
            }), 500

    @app.route('/api/excel/create', methods=['POST'])
    def create_excel_database():
        """Create a new Excel database with the provided room data"""
        try:
            data = request.json
            name = data.get('name', 'New Database')
            rooms = data.get('rooms', [])
            activate = data.get('activate', False)
            
            if not rooms:
                return jsonify({
                    "success": False,
                    "error": "No rooms provided"
                }), 400
                
            # Create the Excel file using the excel creator utility
            from utils.excel_creator import create_excel_database
            
            try:
                file_path = create_excel_database(name, rooms)
                
                # Add to file manager
                with open(file_path, 'rb') as f:
                    from werkzeug.datastructures import FileStorage
                    file = FileStorage(
                        stream=f,
                        filename=f"{name.replace(' ', '_')}.xlsx",
                        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                    file_metadata = service_components.file_manager.add_file(file, title=name)
                
                # Activate if requested
                if activate and file_metadata:
                    service_components.file_manager.set_active_file(file_metadata['id'])
                    service_components.reload()
                
                return jsonify({
                    "success": True,
                    "message": "Database created successfully",
                    "file": file_metadata
                })
            except Exception as e:
                print(f"Error creating database: {str(e)}")
                return jsonify({
                    "success": False,
                    "error": f"Error creating Excel file: {str(e)}"
                }), 500
                
        except Exception as e:
            return jsonify({
                "success": False,
                "error": str(e)
            }), 500
    
    return app