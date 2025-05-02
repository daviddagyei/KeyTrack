import os
import json
import shutil
from datetime import datetime
import uuid
from werkzeug.utils import secure_filename

class ExcelFileManager:
    """
    Manages Excel files for the KeyTrack application:
    - Tracks uploaded files
    - Stores metadata
    - Provides access to file history
    """
    
    def __init__(self, app_root):
        """Initialize with the application root directory"""
        self.app_root = app_root
        self.files_dir = os.path.join(app_root, 'excel_files')
        self.metadata_file = os.path.join(self.files_dir, 'metadata.json')
        
        # Create the directory if it doesn't exist
        os.makedirs(self.files_dir, exist_ok=True)
        
        # Initialize metadata if it doesn't exist
        if not os.path.exists(self.metadata_file):
            self._save_metadata({
                'active_file': None,
                'files': []
            })
    
    def get_metadata(self):
        """Get metadata about all Excel files"""
        if not os.path.exists(self.metadata_file):
            return {
                'active_file': None,
                'files': []
            }
        
        try:
            with open(self.metadata_file, 'r') as f:
                return json.load(f)
        except Exception:
            # If the file is corrupt, return empty metadata
            return {
                'active_file': None,
                'files': []
            }
    
    def _save_metadata(self, metadata):
        """Save metadata to file"""
        with open(self.metadata_file, 'w') as f:
            json.dump(metadata, f, indent=2)
    
    def add_file(self, uploaded_file, title=None):
        """
        Add a new Excel file to the tracked files
        
        Returns:
            dict: File metadata
        """
        # Generate a unique filename
        unique_id = str(uuid.uuid4())
        original_filename = secure_filename(uploaded_file.filename)
        file_id = f"{unique_id}__{original_filename}"
        
        # Save the file to the directory
        file_path = os.path.join(self.files_dir, file_id)
        uploaded_file.save(file_path)
        
        # Create metadata for the file
        file_metadata = {
            'id': file_id,
            'original_name': original_filename,
            'title': title or original_filename,
            'upload_date': datetime.now().isoformat(),
            'path': file_path
        }
        
        # Update metadata
        metadata = self.get_metadata()
        metadata['files'].append(file_metadata)
        self._save_metadata(metadata)
        
        return file_metadata
    
    def set_active_file(self, file_id):
        """
        Set the active Excel file
        
        Args:
            file_id: The ID of the file to set as active
            
        Returns:
            bool: True if successful, False otherwise
        """
        metadata = self.get_metadata()
        
        # Find the file
        file_entry = None
        for entry in metadata['files']:
            if entry['id'] == file_id:
                file_entry = entry
                break
        
        if not file_entry:
            return False
        
        # Copy the file to the active location
        source_path = file_entry['path']
        target_path = os.path.join(self.app_root, 'key_distribution.xlsx')
        
        # Create a backup first
        if os.path.exists(target_path):
            backup_path = os.path.join(self.app_root, 'key_distribution.backup.xlsx')
            shutil.copy2(target_path, backup_path)
        
        # Copy the new file to the target location
        shutil.copy2(source_path, target_path)
        
        # Update the active file in metadata
        metadata['active_file'] = file_id
        self._save_metadata(metadata)
        
        return True
    
    def get_file_path(self, file_id):
        """Get the path to a file by ID"""
        metadata = self.get_metadata()
        
        for entry in metadata['files']:
            if entry['id'] == file_id:
                return entry['path']
        
        return None
    
    def remove_file(self, file_id):
        """
        Remove a file from tracking
        
        Args:
            file_id: The ID of the file to remove
            
        Returns:
            bool: True if successful, False otherwise
        """
        metadata = self.get_metadata()
        
        # Find the file
        file_entry = None
        for idx, entry in enumerate(metadata['files']):
            if entry['id'] == file_id:
                file_entry = entry
                file_index = idx
                break
        
        if not file_entry:
            return False
        
        # Remove the file if it exists
        file_path = file_entry['path']
        if os.path.exists(file_path):
            os.remove(file_path)
        
        # Update metadata
        metadata['files'].pop(file_index)
        
        # If this was the active file, unset it
        if metadata['active_file'] == file_id:
            metadata['active_file'] = None
        
        self._save_metadata(metadata)
        return True
    
    def get_active_file(self):
        """Get the currently active file metadata"""
        metadata = self.get_metadata()
        active_id = metadata['active_file']
        
        if not active_id:
            return None
        
        for entry in metadata['files']:
            if entry['id'] == active_id:
                return entry
        
        return None
    
    def initialize_with_default(self):
        """
        Initialize the system with the default Excel file if it exists
        and no files are currently tracked
        """
        metadata = self.get_metadata()
        
        # If we already have files, don't initialize
        if metadata['files']:
            return False
        
        # Check if the default Excel file exists
        default_path = os.path.join(self.app_root, 'key_distribution.xlsx')
        if not os.path.exists(default_path):
            return False
        
        # Add the default file to our tracking
        unique_id = str(uuid.uuid4())
        file_id = f"{unique_id}__key_distribution.xlsx"
        file_path = os.path.join(self.files_dir, file_id)
        
        # Copy the file
        shutil.copy2(default_path, file_path)
        
        # Create metadata
        file_metadata = {
            'id': file_id,
            'original_name': 'key_distribution.xlsx',
            'title': 'Default Excel File',
            'upload_date': datetime.now().isoformat(),
            'path': file_path
        }
        
        # Update metadata
        metadata['files'].append(file_metadata)
        metadata['active_file'] = file_id
        self._save_metadata(metadata)
        
        return True