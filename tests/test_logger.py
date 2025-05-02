import pytest
import logging
import os
import tempfile
import time
from unittest.mock import patch, MagicMock, call
from pathlib import Path

from utils import logger


class TestLogger:
    """Test cases for the logger utility."""
    
    def test_logger_exists(self):
        """Test that the logger object exists."""
        assert hasattr(logger, 'logger')
        assert isinstance(logger.logger, logging.Logger)
    
    def test_logger_name(self):
        """Test that the logger has the correct name."""
        assert logger.logger.name == "keytrack"
    
    def test_logger_level(self):
        """Test that the logger has the correct level."""
        assert logger.logger.level == logging.INFO
    
    def test_convenience_functions_exist(self):
        """Test that all convenience functions exist."""
        assert hasattr(logger, 'info')
        assert hasattr(logger, 'error')
        assert hasattr(logger, 'warning')
        assert hasattr(logger, 'debug')
        
        # Check that they are callable
        assert callable(logger.info)
        assert callable(logger.error)
        assert callable(logger.warning)
        assert callable(logger.debug)
    
    @patch('utils.logger.logger')
    def test_info_function_calls_logger(self, mock_logger):
        """Test that the info function calls the logger's info method."""
        message = "Test info message"
        logger.info(message)
        mock_logger.info.assert_called_once_with(message)
    
    @patch('utils.logger.logger')
    def test_info_function_with_complex_message(self, mock_logger):
        """Test the info function with a more complex message."""
        message = "User 'John Doe' collected key for room 101 at 2025-05-01 10:30:00"
        logger.info(message)
        mock_logger.info.assert_called_once_with(message)
    
    @patch('utils.logger.logger')
    def test_error_function_calls_logger(self, mock_logger):
        """Test that the error function calls the logger's error method."""
        message = "Test error message"
        logger.error(message)
        mock_logger.error.assert_called_once_with(message)
    
    @patch('utils.logger.logger')
    def test_error_function_with_exception_message(self, mock_logger):
        """Test the error function with an exception message."""
        try:
            raise ValueError("Something went wrong")
        except ValueError as e:
            message = f"Caught an exception: {str(e)}"
            logger.error(message)
            mock_logger.error.assert_called_once_with(message)
    
    @patch('utils.logger.logger')
    def test_warning_function_calls_logger(self, mock_logger):
        """Test that the warning function calls the logger's warning method."""
        message = "Test warning message"
        logger.warning(message)
        mock_logger.warning.assert_called_once_with(message)
    
    @patch('utils.logger.logger')
    def test_warning_function_with_complex_message(self, mock_logger):
        """Test the warning function with a more complex message."""
        message = "Room 101 has only 1 key remaining"
        logger.warning(message)
        mock_logger.warning.assert_called_once_with(message)
    
    @patch('utils.logger.logger')
    def test_debug_function_calls_logger(self, mock_logger):
        """Test that the debug function calls the logger's debug method."""
        message = "Test debug message"
        logger.debug(message)
        mock_logger.debug.assert_called_once_with(message)
    
    @patch('utils.logger.logger')
    def test_debug_function_with_complex_message(self, mock_logger):
        """Test the debug function with a more complex message."""
        message = "Processing key action: KeyAction(Student='John Doe', timestamp='2025-05-01 10:30:00')"
        logger.debug(message)
        mock_logger.debug.assert_called_once_with(message)
    
    def test_log_directory_creation(self):
        """Test that the logs directory is created."""
        # Check that the logs directory exists
        assert os.path.exists("logs")
        assert os.path.isdir("logs")
    
    def test_setup_logger_returns_logger(self):
        """Test that setup_logger returns a logger object."""
        # Reset the logger to force re-initialization
        logger.logger = None
        
        # Recreate the logger
        new_logger = logger.setup_logger()
        
        assert isinstance(new_logger, logging.Logger)
        assert new_logger.name == "keytrack"
    
    def test_setup_logger_creates_file_handler(self):
        """Test that setup_logger creates a file handler."""
        # Reset the logger to force re-initialization
        logger.logger = None
        
        # Recreate the logger
        new_logger = logger.setup_logger()
        
        # Check that there are handlers
        assert len(new_logger.handlers) > 0
        
        # Check that at least one handler is a FileHandler
        assert any(isinstance(handler, logging.FileHandler) for handler in new_logger.handlers)
    
    def test_setup_logger_creates_console_handler(self):
        """Test that setup_logger creates a console handler."""
        # Reset the logger to force re-initialization
        logger.logger = None
        
        # Recreate the logger
        new_logger = logger.setup_logger()
        
        # Check that there are handlers
        assert len(new_logger.handlers) > 0
        
        # Check that at least one handler is a StreamHandler
        assert any(isinstance(handler, logging.StreamHandler) and not isinstance(handler, logging.FileHandler) 
                  for handler in new_logger.handlers)
    
    @patch('logging.Logger.addHandler')
    @patch('logging.FileHandler')
    @patch('logging.StreamHandler')
    def test_setup_logger_handlers_and_formatters(self, mock_stream_handler, mock_file_handler, mock_add_handler):
        """Test that the logger setup adds the correct handlers and formatters."""
        # Reset the logger to force re-initialization
        logger.logger = None
        
        # Mock the handlers
        mock_file_handler_instance = MagicMock()
        mock_stream_handler_instance = MagicMock()
        mock_file_handler.return_value = mock_file_handler_instance
        mock_stream_handler.return_value = mock_stream_handler_instance
        
        # Recreate the logger
        new_logger = logger.setup_logger()
        
        # Check that both handlers were added
        assert mock_add_handler.call_count == 2
        
        # Check that formatters were set for both handlers
        mock_file_handler_instance.setFormatter.assert_called_once()
        mock_stream_handler_instance.setFormatter.assert_called_once()
    
    def test_logging_to_temp_file(self):
        """Test that logs are written to a file."""
        # Create a temporary log file
        temp_log_file = tempfile.NamedTemporaryFile(delete=False)
        temp_log_file.close()
        
        try:
            # Create a custom logger that writes to our temp file
            test_logger = logging.getLogger("test_keytrack")
            test_logger.setLevel(logging.INFO)
            
            # Add a file handler that points to our temp file
            handler = logging.FileHandler(temp_log_file.name)
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            test_logger.addHandler(handler)
            
            # Log some messages
            test_logger.info("Test info message to file")
            test_logger.error("Test error message to file")
            
            # Close the handler to ensure all data is written
            handler.close()
            
            # Check the contents of the log file
            with open(temp_log_file.name, 'r') as f:
                log_contents = f.read()
                assert "Test info message to file" in log_contents
                assert "Test error message to file" in log_contents
        
        finally:
            # Clean up
            os.unlink(temp_log_file.name)