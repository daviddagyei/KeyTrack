class KeyTrackException(Exception):
    """Base exception for all KeyTrack application errors."""
    pass


class RoomNotFoundException(KeyTrackException):
    """Raised when a specified room cannot be found."""
    def __init__(self, room_id):
        self.room_id = room_id
        super().__init__(f"Room {room_id} not found.")


class NoKeysAvailableException(KeyTrackException):
    """Raised when no keys are available for a room."""
    def __init__(self, room_id):
        self.room_id = room_id
        super().__init__(f"No keys available for room {room_id}.")


class KeyNotCollectedException(KeyTrackException):
    """Raised when trying to return a key that was not collected."""
    def __init__(self, room_id, student_name):
        self.room_id = room_id
        self.student_name = student_name
        super().__init__(f"No record of key collection for room {room_id} by student {student_name}.")


class DataAccessException(KeyTrackException):
    """Raised when there's an error accessing the data storage."""
    pass