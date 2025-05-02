import sys
import os
import datetime
# Add the parent directory to sys.path to allow importing backend module
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from flask import Flask, render_template
from backend.server import create_app

def create_webapp():
    """Create the webapp Flask application"""
    app = Flask(__name__)
    
    # Add context processor to provide current date to templates
    @app.context_processor
    def inject_now():
        return {'now': datetime.datetime.now()}
    
    # Routes for the web interface
    @app.route('/')
    def index():
        """Render the main dashboard page"""
        return render_template('index.html')
    
    @app.route('/room/<room_id>')
    def room_detail(room_id):
        """Render the room detail page"""
        # Ensure the room_id is decoded for display
        decoded_room_id = room_id.replace('%20', ' ')
        return render_template('room_detail.html', room_id=decoded_room_id)
    
    @app.route('/actions')
    def actions():
        """Render the key actions page"""
        return render_template('actions.html')
    
    @app.route('/student')
    def student():
        """Render the student lookup page"""
        return render_template('student.html')
        
    @app.route('/excel')
    def excel_management():
        """Render the Excel management page"""
        return render_template('excel_management.html')
        
    @app.route('/create-database')
    def create_database():
        """Render the database creation page"""
        return render_template('create_database.html')
    
    # Mount the API as a blueprint
    api_app = create_app()
    
    # Register all routes from the API app, skipping any conflicting endpoints
    for rule in api_app.url_map.iter_rules():
        if rule.endpoint != 'static':  # Skip the static endpoint to avoid conflict
            endpoint = api_app.view_functions[rule.endpoint]
            app.add_url_rule(rule.rule, rule.endpoint, endpoint, methods=rule.methods)
    
    return app

if __name__ == '__main__':
    app = create_webapp()
    app.run(debug=True, port=5000)