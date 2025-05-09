import sys
import os
import datetime
import secrets
# Add the parent directory to sys.path to allow importing backend module
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from flask import Flask, render_template, request, redirect, url_for, flash, session
from backend.server import create_app
from utils.auth import sign_up, sign_in, sign_out, set_user_session, clear_user_session, enable_demo_mode, login_required

def create_webapp():
    """Create the webapp Flask application"""
    app = Flask(__name__)
    
    # Configure the Flask application with a secret key for sessions
    app.secret_key = os.environ.get('FLASK_SECRET_KEY', secrets.token_hex(16))
    
    # Add context processor to provide current date to templates
    @app.context_processor
    def inject_now():
        return {'now': datetime.datetime.now()}
    
    # Authentication routes
    @app.route('/login', methods=['GET', 'POST'])
    def login():
        """Handle user login"""
        if request.method == 'POST':
            email = request.form.get('email')
            password = request.form.get('password')
            
            user, session_data, error = sign_in(email, password)
            
            if error:
                flash(f'Login failed: {error}', 'danger')
                return render_template('login.html')
            
            set_user_session(user, session_data)
            flash('You have successfully logged in!', 'success')
            return redirect(url_for('index'))
        
        return render_template('login.html')
    
    @app.route('/signup', methods=['GET', 'POST'])
    def signup():
        """Handle user registration"""
        if request.method == 'POST':
            email = request.form.get('email')
            password = request.form.get('password')
            confirm_password = request.form.get('confirm_password')
            name = request.form.get('name', '')
            
            # Basic validation
            if password != confirm_password:
                flash('Passwords do not match', 'danger')
                return render_template('signup.html')
            
            if len(password) < 6:
                flash('Password must be at least 6 characters long', 'danger')
                return render_template('signup.html')
            
            # Create user with Supabase
            user_metadata = {'name': name} if name else None
            user, error = sign_up(email, password, user_metadata)
            
            if error:
                flash(f'Registration failed: {error}', 'danger')
                return render_template('signup.html')
            
            flash('Account created! Please check your email to confirm your registration.', 'success')
            return redirect(url_for('login'))
        
        return render_template('signup.html')
    
    @app.route('/logout')
    def logout():
        """Handle user logout"""
        success, error = sign_out()
        clear_user_session()
        session.pop('demo_mode', None)  # Also clear demo mode if active
        
        if error:
            flash(f'Logout error: {error}', 'warning')
        else:
            flash('You have been logged out successfully', 'success')
            
        return redirect(url_for('login'))
    
    @app.route('/demo')
    def demo_mode():
        """Enable demo mode without authentication"""
        enable_demo_mode()
        flash('Demo mode activated! You can explore all features without logging in.', 'info')
        return redirect(url_for('index'))
    
    # Routes for the web interface with auth protection
    @app.route('/')
    @login_required
    def index():
        """Render the main dashboard page"""
        return render_template('index.html')
    
    @app.route('/room/<room_id>')
    @login_required
    def room_detail(room_id):
        """Render the room detail page"""
        # Ensure the room_id is decoded for display
        decoded_room_id = room_id.replace('%20', ' ')
        return render_template('room_detail.html', room_id=decoded_room_id)
    
    @app.route('/actions')
    @login_required
    def actions():
        """Render the key actions page"""
        return render_template('actions.html')
    
    @app.route('/student')
    @login_required
    def student():
        """Render the student lookup page"""
        return render_template('student.html')
        
    @app.route('/excel')
    @login_required
    def excel_management():
        """Render the Excel management page"""
        return render_template('excel_management.html')
        
    @app.route('/create-database')
    @login_required
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
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
        debug=True
    )