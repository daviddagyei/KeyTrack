// KeyTrack Room Detail JavaScript
document.addEventListener('DOMContentLoaded', function() {
    // Get room ID from URL path and properly decode it
    const pathParts = window.location.pathname.split('/');
    const roomId = decodeURIComponent(pathParts[pathParts.length - 1]);
    
    // Update room ID in the UI
    document.getElementById('roomId').textContent = roomId;
    document.getElementById('roomIdBreadcrumb').textContent = roomId;
    
    // Load room data
    loadRoomData(roomId);
    
    // Set up action buttons
    setupActionButtons(roomId);
});

// Load room data from the API
function loadRoomData(roomId) {
    fetch(`/api/rooms/${roomId}`)
        .then(response => {
            if (!response.ok) {
                throw new Error(`Failed to load room data (${response.status})`);
            }
            return response.json();
        })
        .then(data => {
            if (data.success) {
                displayRoomData(data.room);
                hideLoadingIndicator();
            } else {
                showError(data.error || 'Failed to load room data');
            }
        })
        .catch(error => {
            console.error('Error loading room data:', error);
            showError(error.message || 'Failed to load room data');
        });
}

// Display room data in the UI
function displayRoomData(room) {
    // Update key counts
    document.getElementById('totalKeys').textContent = room.total_keys;
    document.getElementById('availableKeys').textContent = room.available_keys;
    document.getElementById('collectedKeys').textContent = room.collected_keys;
    document.getElementById('lostKeys').textContent = room.lost_keys;
    
    // Display collected keys
    const collectedKeysList = document.getElementById('collectedKeysList');
    const noCollectedKeys = document.getElementById('noCollectedKeys');
    
    if (room.collected_actions && room.collected_actions.length > 0) {
        collectedKeysList.innerHTML = '';
        room.collected_actions.forEach(action => {
            collectedKeysList.appendChild(createActionListItem(action, 'collected'));
        });
        collectedKeysList.classList.remove('d-none');
        noCollectedKeys.classList.add('d-none');
    } else {
        collectedKeysList.classList.add('d-none');
        noCollectedKeys.classList.remove('d-none');
    }
    
    // Display returned keys
    const returnedKeysList = document.getElementById('returnedKeysList');
    const noReturnedKeys = document.getElementById('noReturnedKeys');
    
    if (room.returned_actions && room.returned_actions.length > 0) {
        returnedKeysList.innerHTML = '';
        room.returned_actions.forEach(action => {
            returnedKeysList.appendChild(createActionListItem(action, 'returned'));
        });
        returnedKeysList.classList.remove('d-none');
        noReturnedKeys.classList.add('d-none');
    } else {
        returnedKeysList.classList.add('d-none');
        noReturnedKeys.classList.remove('d-none');
    }
    
    // Display lost keys
    const lostKeysList = document.getElementById('lostKeysList');
    const noLostKeys = document.getElementById('noLostKeys');
    
    if (room.lost_actions && room.lost_actions.length > 0) {
        lostKeysList.innerHTML = '';
        room.lost_actions.forEach(action => {
            lostKeysList.appendChild(createActionListItem(action, 'lost'));
        });
        lostKeysList.classList.remove('d-none');
        noLostKeys.classList.add('d-none');
    } else {
        lostKeysList.classList.add('d-none');
        noLostKeys.classList.remove('d-none');
    }
    
    // Display borrowed keys
    const borrowedKeysList = document.getElementById('borrowedKeysList');
    const noBorrowedKeys = document.getElementById('noBorrowedKeys');
    
    if (room.borrowed_actions && room.borrowed_actions.length > 0) {
        borrowedKeysList.innerHTML = '';
        room.borrowed_actions.forEach(action => {
            borrowedKeysList.appendChild(createActionListItem(action, 'borrowed'));
        });
        borrowedKeysList.classList.remove('d-none');
        noBorrowedKeys.classList.add('d-none');
    } else {
        borrowedKeysList.classList.add('d-none');
        noBorrowedKeys.classList.remove('d-none');
    }
}

// Create a list item for a key action
function createActionListItem(action, type) {
    const listItem = document.createElement('div');
    listItem.className = 'list-group-item d-flex justify-content-between align-items-center';
    
    let badgeClass = 'bg-maroon';
    let displayType = type;
    
    // Use consistent styling for different action types
    if (type === 'returned') {
        badgeClass = 'bg-success';
        // Check if the returned key was originally collected or borrowed
        // This is a visual enhancement only - backend now treats all returns the same
        const returnDate = new Date(action.timestamp);
        const collectedActions = document.querySelectorAll('#collectedKeysList .list-group-item');
        const borrowedActions = document.querySelectorAll('#borrowedKeysList .list-group-item');
        
        let wasCollected = false;
        let wasBorrowed = false;
        
        // Check if this student collected a key before returning it
        for (let i = 0; i < collectedActions.length; i++) {
            const actionStudent = collectedActions[i].querySelector('strong').textContent;
            if (actionStudent === action.student) {
                wasCollected = true;
                break;
            }
        }
        
        // If not collected, check if it was borrowed
        if (!wasCollected) {
            for (let i = 0; i < borrowedActions.length; i++) {
                const actionStudent = borrowedActions[i].querySelector('strong').textContent;
                if (actionStudent === action.student) {
                    wasBorrowed = true;
                    break;
                }
            }
        }
        
        // Adjust display type based on collection history
        if (wasBorrowed) {
            displayType = 'returned borrowed';
        } else {
            displayType = 'returned collected';
        }
    }
    if (type === 'lost') badgeClass = 'bg-danger';
    if (type === 'borrowed') badgeClass = 'bg-warning';
    
    listItem.innerHTML = `
        <div>
            <strong>${action.student}</strong>
        </div>
        <div>
            <span class="badge ${badgeClass}">${displayType}</span>
            <small class="text-muted ms-2">${new Date(action.timestamp).toLocaleString()}</small>
        </div>
    `;
    
    return listItem;
}

// Hide the loading indicator and show the room details
function hideLoadingIndicator() {
    document.getElementById('roomLoading').classList.add('d-none');
    document.getElementById('roomDetails').classList.remove('d-none');
}

// Show an error message
function showError(message) {
    document.getElementById('roomLoading').classList.add('d-none');
    
    const errorModal = new bootstrap.Modal(document.getElementById('errorModal'));
    document.getElementById('errorModalText').textContent = message;
    errorModal.show();
}

// Show a success message
function showSuccess(message) {
    const successModal = new bootstrap.Modal(document.getElementById('successModal'));
    document.getElementById('successModalText').textContent = message;
    successModal.show();
    
    // Reload room data after a successful action, but keep the modal visible
    setTimeout(() => {
        const pathParts = window.location.pathname.split('/');
        const roomId = decodeURIComponent(pathParts[pathParts.length - 1]);
        loadRoomData(roomId);
        
        // Close the modal after another delay to ensure users can read the message
        setTimeout(() => {
            successModal.hide();
        }, 3000); // Keep success message visible for 3 seconds after data reload
    }, 1000); // Wait 1 second before reloading data
}

// Set up action buttons (collect, return, lost, borrow)
function setupActionButtons(roomId) {
    // Collect key
    document.getElementById('collectKeyButton').addEventListener('click', function() {
        const studentName = document.getElementById('collectStudentName').value.trim();
        if (!studentName) {
            document.getElementById('collectKeyError').textContent = 'Please enter a student name';
            document.getElementById('collectKeyError').classList.remove('d-none');
            return;
        }
        
        performAction('collect', roomId, studentName);
    });
    
    // Return key
    document.getElementById('returnKeyButton').addEventListener('click', function() {
        const studentName = document.getElementById('returnStudentName').value.trim();
        if (!studentName) {
            document.getElementById('returnKeyError').textContent = 'Please enter a student name';
            document.getElementById('returnKeyError').classList.remove('d-none');
            return;
        }
        
        performAction('return', roomId, studentName);
    });
    
    // Report lost key
    document.getElementById('lostKeyButton').addEventListener('click', function() {
        const studentName = document.getElementById('lostStudentName').value.trim();
        if (!studentName) {
            document.getElementById('lostKeyError').textContent = 'Please enter a student name';
            document.getElementById('lostKeyError').classList.remove('d-none');
            return;
        }
        
        performAction('lost', roomId, studentName);
    });
    
    // Borrow spare key
    document.getElementById('borrowKeyButton').addEventListener('click', function() {
        const studentName = document.getElementById('borrowStudentName').value.trim();
        if (!studentName) {
            document.getElementById('borrowKeyError').textContent = 'Please enter a student name';
            document.getElementById('borrowKeyError').classList.remove('d-none');
            return;
        }
        
        performAction('borrow', roomId, studentName);
    });
    
    // Return borrowed key
    document.getElementById('returnBorrowedKeyButton').addEventListener('click', function() {
        const studentName = document.getElementById('returnBorrowedStudentName').value.trim();
        if (!studentName) {
            document.getElementById('returnBorrowedKeyError').textContent = 'Please enter a student name';
            document.getElementById('returnBorrowedKeyError').classList.remove('d-none');
            return;
        }
        
        performAction('returnBorrowed', roomId, studentName);
    });
}

// Perform a key action (collect, return, lost, borrow, returnBorrowed)
function performAction(action, roomId, studentName) {
    // Map action to API endpoint
    const actionEndpoints = {
        'collect': 'collect',
        'return': 'return',
        'lost': 'lost',
        'borrow': 'borrow',
        'returnBorrowed': 'return-borrowed'
    };
    
    const endpoint = actionEndpoints[action];
    
    // Hide previous error messages
    document.querySelectorAll('.alert-danger').forEach(el => {
        el.classList.add('d-none');
    });
    
    // Send request to API
    fetch(`/api/actions/${endpoint}`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            room_id: roomId,
            student_name: studentName
        })
    })
    .then(response => response.json())
    .then(data => {
        // Show success or error message
        if (data.success) {
            // Only close the modal and clear input on success
            const modalId = `${action}KeyModal`;
            const modal = bootstrap.Modal.getInstance(document.getElementById(modalId));
            modal.hide();
            
            // Clear the input field
            document.getElementById(`${action}StudentName`).value = '';
            
            // Show success message
            showSuccess(data.message);
        } else {
            // Show error inside the modal (don't close the modal)
            const errorModalId = `${action}KeyError`;
            document.getElementById(errorModalId).textContent = data.error || 'An error occurred';
            document.getElementById(errorModalId).classList.remove('d-none');
        }
    })
    .catch(error => {
        console.error(`Error performing ${action} action:`, error);
        const errorModalId = `${action}KeyError`;
        document.getElementById(errorModalId).textContent = 'Failed to perform action. Please try again.';
        document.getElementById(errorModalId).classList.remove('d-none');
    });
}