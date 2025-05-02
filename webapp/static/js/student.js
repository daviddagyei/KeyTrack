document.addEventListener('DOMContentLoaded', function() {
    const studentSearchForm = document.getElementById('studentSearchForm');
    const studentNameInput = document.getElementById('studentName');
    const studentResults = document.getElementById('studentResults');
    const studentLoading = document.getElementById('studentLoading');
    const studentNameDisplay = document.getElementById('studentNameDisplay');
    const studentHistoryTable = document.getElementById('studentHistoryTable');
    const noResults = document.getElementById('noResults');
    
    studentSearchForm.addEventListener('submit', function(e) {
        e.preventDefault();
        
        const studentName = studentNameInput.value.trim();
        if (!studentName) return;
        
        // Show loading spinner
        studentResults.classList.add('d-none');
        studentLoading.classList.remove('d-none');
        
        // Fetch student history from API
        fetch(`/api/student/${encodeURIComponent(studentName)}`)
            .then(response => response.json())
            .then(data => {
                // Hide loading spinner
                studentLoading.classList.add('d-none');
                
                if (data.success) {
                    // Update student name in results
                    studentNameDisplay.textContent = data.student_name;
                    
                    // Clear previous results
                    studentHistoryTable.innerHTML = '';
                    
                    // Check if there's any history
                    if (data.history && data.history.length > 0) {
                        // Populate table with student history
                        data.history.forEach(item => {
                            const row = document.createElement('tr');
                            
                            // Format the action type for display
                            let actionType = item.action_type;
                            actionType = actionType.charAt(0).toUpperCase() + actionType.slice(1);
                            
                            // Format timestamp
                            const timestamp = new Date(item.timestamp);
                            const formattedDate = timestamp.toLocaleDateString();
                            const formattedTime = timestamp.toLocaleTimeString();
                            
                            row.innerHTML = `
                                <td>${item.room_id}</td>
                                <td>${actionType}</td>
                                <td>${formattedDate} ${formattedTime}</td>
                            `;
                            
                            studentHistoryTable.appendChild(row);
                        });
                        
                        // Show results, hide no results message
                        studentResults.classList.remove('d-none');
                        noResults.classList.add('d-none');
                    } else {
                        // Show no results message
                        studentResults.classList.remove('d-none');
                        noResults.classList.remove('d-none');
                    }
                } else {
                    // Show error message
                    alert('Error: ' + (data.error || 'Failed to fetch student history'));
                }
            })
            .catch(error => {
                // Hide loading spinner
                studentLoading.classList.add('d-none');
                
                // Show error message
                console.error('Error fetching student history:', error);
                alert('Error: Failed to fetch student history');
            });
    });
});