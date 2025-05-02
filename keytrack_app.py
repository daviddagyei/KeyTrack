from flask import Flask, send_from_directory, render_template
from backend.server import create_app
import os

# Create the Flask application
app = create_app()

# Serve static files
@app.route('/static/<path:path>')
def serve_static(path):
    return send_from_directory('webapp/static', path)

# Serve the UI
@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/rooms/<room_id>', methods=['GET'])
def room_detail(room_id):
    return render_template('room_detail.html')

@app.route('/student', methods=['GET'])
def student():
    return render_template('student.html')

@app.route('/actions', methods=['GET'])
def actions():
    return render_template('actions.html')

if __name__ == '__main__':
    app.template_folder = 'webapp/templates'
    app.static_folder = 'webapp/static'
    app.run(debug=True, port=5000)