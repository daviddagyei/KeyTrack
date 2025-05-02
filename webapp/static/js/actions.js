document.addEventListener('DOMContentLoaded', function() {
    // Get references to all forms
    const collectKeyForm = document.getElementById('collectKeyForm');
    const returnKeyForm = document.getElementById('returnKeyForm');
    const lostKeyForm = document.getElementById('lostKeyForm');
    const borrowKeyForm = document.getElementById('borrowKeyForm');
    
    // Success modal elements
    const successModal = new bootstrap.Modal(document.getElementById('successModal'));
    const successModalText = document.getElementById('successModalText');
    
    // Handle collect key form submission
    collectKeyForm.addEventListener('submit', function(e) {
        e.preventDefault();
        const roomId = document.getElementById('collectRoomId').value;
        const studentName = document.getElementById('collectStudentName').value;
        const errorDiv = document.getElementById('collectKeyError');
        
        fetch('/api/actions/collect', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                room_id: roomId,
                student_name: studentName
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                // Show success message
                successModalText.textContent = data.message;
                successModal.show();
                // Reset form
                collectKeyForm.reset();
                errorDiv.classList.add('d-none');
            } else {
                // Show error message
                errorDiv.textContent = data.error;
                errorDiv.classList.remove('d-none');
            }
        })
        .catch(error => {
            errorDiv.textContent = 'An error occurred while processing your request.';
            errorDiv.classList.remove('d-none');
            console.error('Error:', error);
        });
    });
    
    // Handle return key form submission
    returnKeyForm.addEventListener('submit', function(e) {
        e.preventDefault();
        const roomId = document.getElementById('returnRoomId').value;
        const studentName = document.getElementById('returnStudentName').value;
        const errorDiv = document.getElementById('returnKeyError');
        
        fetch('/api/actions/return', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                room_id: roomId,
                student_name: studentName
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                // Show success message
                successModalText.textContent = data.message;
                successModal.show();
                // Reset form
                returnKeyForm.reset();
                errorDiv.classList.add('d-none');
            } else {
                // Show error message
                errorDiv.textContent = data.error;
                errorDiv.classList.remove('d-none');
            }
        })
        .catch(error => {
            errorDiv.textContent = 'An error occurred while processing your request.';
            errorDiv.classList.remove('d-none');
            console.error('Error:', error);
        });
    });
    
    // Handle report lost key form submission
    lostKeyForm.addEventListener('submit', function(e) {
        e.preventDefault();
        const roomId = document.getElementById('lostRoomId').value;
        const studentName = document.getElementById('lostStudentName').value;
        const errorDiv = document.getElementById('lostKeyError');
        
        fetch('/api/actions/lost', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                room_id: roomId,
                student_name: studentName
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                // Show success message
                successModalText.textContent = data.message;
                successModal.show();
                // Reset form
                lostKeyForm.reset();
                errorDiv.classList.add('d-none');
            } else {
                // Show error message
                errorDiv.textContent = data.error;
                errorDiv.classList.remove('d-none');
            }
        })
        .catch(error => {
            errorDiv.textContent = 'An error occurred while processing your request.';
            errorDiv.classList.remove('d-none');
            console.error('Error:', error);
        });
    });
    
    // Handle borrow spare key form submission
    borrowKeyForm.addEventListener('submit', function(e) {
        e.preventDefault();
        const roomId = document.getElementById('borrowRoomId').value;
        const studentName = document.getElementById('borrowStudentName').value;
        const errorDiv = document.getElementById('borrowKeyError');
        
        fetch('/api/actions/borrow', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                room_id: roomId,
                student_name: studentName
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                // Show success message
                successModalText.textContent = data.message;
                successModal.show();
                // Reset form
                borrowKeyForm.reset();
                errorDiv.classList.add('d-none');
            } else {
                // Show error message
                errorDiv.textContent = data.error;
                errorDiv.classList.remove('d-none');
            }
        })
        .catch(error => {
            errorDiv.textContent = 'An error occurred while processing your request.';
            errorDiv.classList.remove('d-none');
            console.error('Error:', error);
        });
    });
});