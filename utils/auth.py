import os
import re
from functools import wraps
from supabase import create_client, Client
from flask import session, redirect, url_for, flash
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Initialize Supabase client
supabase_url = os.environ.get("SUPABASE_URL")
supabase_key = os.environ.get("SUPABASE_KEY")

supabase: Client = None

def init_supabase():
    """Initialize Supabase client"""
    global supabase
    if not supabase:
        if supabase_url and supabase_key:
            supabase = create_client(supabase_url, supabase_key)
        else:
            print("Warning: Supabase credentials not found. Auth features will be disabled.")
    return supabase

def is_uchicago_email(email):
    """Verify that the email is a UChicago email address"""
    # Simple check for @uchicago.edu domain
    pattern = r'^[a-zA-Z0-9_.+-]+@uchicago\.edu$'
    return bool(re.match(pattern, email))

def sign_up(email, password, user_metadata=None):
    """Register a new user with Supabase"""
    try:
        # Verify UChicago email
        if not is_uchicago_email(email):
            return None, "You must use a valid @uchicago.edu email address"
            
        client = init_supabase()
        if not client:
            return None, "Supabase authentication is not configured"
            
        response = client.auth.sign_up({
            "email": email,
            "password": password,
            "options": {
                "data": user_metadata or {}
            }
        })
        return response.user, None
    except Exception as e:
        return None, str(e)

def sign_in(email, password):
    """Sign in an existing user"""
    try:
        # Verify UChicago email
        if not is_uchicago_email(email):
            return None, None, "You must use a valid @uchicago.edu email address"
            
        client = init_supabase()
        if not client:
            return None, "Supabase authentication is not configured"
            
        response = client.auth.sign_in_with_password({
            "email": email,
            "password": password
        })
        return response.user, response.session, None
    except Exception as e:
        return None, None, str(e)

def sign_out():
    """Sign out the current user"""
    try:
        client = init_supabase()
        if client:
            client.auth.sign_out()
        return True, None
    except Exception as e:
        return False, str(e)

def get_current_user():
    """Get the current authenticated user"""
    try:
        client = init_supabase()
        if not client:
            return None
            
        response = client.auth.get_user()
        return response.user if response else None
    except:
        return None

def refresh_session():
    """Refresh the user's session token"""
    try:
        client = init_supabase()
        if not client:
            return None
            
        response = client.auth.refresh_session()
        return response.session
    except:
        return None

# Flask authentication decorators and helpers
def login_required(f):
    """Decorator to require login for a route, allows demo mode"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # Check if user is in demo mode
        if session.get('demo_mode'):
            return f(*args, **kwargs)
            
        # Check if user is logged in with Supabase
        if not session.get('user'):
            flash('Please log in to access this page', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def set_user_session(user, auth_session):
    """Store user data in Flask session"""
    session['user'] = {
        'id': user.id,
        'email': user.email,
        'user_metadata': user.user_metadata
    }
    # Store access token in session
    if auth_session:
        session['access_token'] = auth_session.access_token
        session['refresh_token'] = auth_session.refresh_token
    
def clear_user_session():
    """Clear user data from Flask session"""
    session.pop('user', None)
    session.pop('access_token', None)
    session.pop('refresh_token', None)
    
def enable_demo_mode():
    """Enable demo mode (no authentication needed)"""
    session['demo_mode'] = True
    
def disable_demo_mode():
    """Disable demo mode"""
    session.pop('demo_mode', None)