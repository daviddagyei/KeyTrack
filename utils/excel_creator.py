import os
import tempfile
import pandas as pd
from datetime import datetime

def create_excel_database(name, rooms_data):
    """
    Create a new Excel database file with the provided room data
    
    Args:
        name: The name of the database
        rooms_data: A list of room data objects with 'id' and 'keys' properties
        
    Returns:
        str: The path to the created Excel file
    """
    # Create temporary file path
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp:
        temp_path = temp.name
    
    # Create Excel writer
    writer = pd.ExcelWriter(temp_path, engine='openpyxl')
    
    # Create Room Data sheet
    room_rows = []
    for room in rooms_data:
        room_id = room.get('id', '')
        total_keys = room.get('keys', 0)
        
        room_rows.append({
            'Room Number': room_id,  # Changed from 'Room ID' to match standard
            'Available Keys': total_keys,  # All keys are initially available
            'Collected By': "",      # Initialize empty string for Collected By
            'Lost Keys': "",         # Initialize empty string for Lost Keys
            'Borrowed Spare Keys': "",  # Initialize empty string for Borrowed Spare Keys
            'Returned Keys': ""      # Initialize empty string for Returned Keys
        })
    
    # Convert to DataFrame and write to Excel
    if room_rows:
        rooms_df = pd.DataFrame(room_rows)
        rooms_df.to_excel(writer, sheet_name='Room Data', index=False)
    
    # Create a History sheet with column headers but no data yet
    history_df = pd.DataFrame(columns=[
        'Room ID', 'Student', 'Action', 'Timestamp'
    ])
    history_df.to_excel(writer, sheet_name='History', index=False)
    
    # Create a metadata sheet with database information
    metadata = {
        'Name': [name],
        'Created Date': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
        'Total Rooms': [len(rooms_data)],
        'Total Keys': [sum(room.get('keys', 0) for room in rooms_data)]
    }
    metadata_df = pd.DataFrame(metadata)
    metadata_df.to_excel(writer, sheet_name='Metadata', index=False)
    
    # Save the Excel file
    writer.close()
    
    # Copy to a final path in the application directory
    app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    safe_name = name.replace(' ', '_').lower() + '__' + datetime.now().strftime('%Y%m%d_%H%M%S')
    final_path = os.path.join(app_root, 'excel_files', f"{safe_name}.xlsx")
    
    # Make sure directory exists
    os.makedirs(os.path.dirname(final_path), exist_ok=True)
    
    # Copy the file
    import shutil
    shutil.copy2(temp_path, final_path)
    
    # Delete the temp file
    os.unlink(temp_path)
    
    return final_path