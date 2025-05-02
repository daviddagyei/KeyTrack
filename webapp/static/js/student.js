document.addEventListener('DOMContentLoaded', function() {
    const studentSearchForm = document.getElementById('studentSearchForm');
    const studentNameInput = document.getElementById('studentName');
    const studentResults = document.getElementById('studentResults');
    const studentLoading = document.getElementById('studentLoading');
    const studentNameDisplay = document.getElementById('studentNameDisplay');
    const studentHistoryTable = document.getElementById('studentHistoryTable');
    const noResults = document.getElementById('noResults');
    const studentSuggestions = document.getElementById('studentSuggestions');
    
    // Variables for suggestion functionality
    let debounceTimeout = null;
    let currentSelection = -1;
    let suggestions = [];
    
    // Add event listener for input to trigger suggestions
    studentNameInput.addEventListener('input', function() {
        const query = this.value.trim();
        
        // Clear previous timeout to prevent multiple requests
        if (debounceTimeout) {
            clearTimeout(debounceTimeout);
        }
        
        // Reset current selection
        currentSelection = -1;
        
        // If the query is empty, hide the suggestions
        if (!query) {
            hideSuggestions();
            return;
        }
        
        // Debounce to avoid too many requests
        debounceTimeout = setTimeout(() => {
            fetchSuggestions(query);
        }, 300);
    });
    
    // Handle keyboard navigation within suggestions
    studentNameInput.addEventListener('keydown', function(e) {
        // Only proceed if suggestions are visible
        if (studentSuggestions.style.display !== 'block') return;
        
        switch(e.key) {
            case 'ArrowDown':
                e.preventDefault();
                if (currentSelection < suggestions.length - 1) {
                    currentSelection++;
                    highlightSuggestion();
                }
                break;
            case 'ArrowUp':
                e.preventDefault();
                if (currentSelection > 0) {
                    currentSelection--;
                    highlightSuggestion();
                }
                break;
            case 'Enter':
                if (currentSelection >= 0 && currentSelection < suggestions.length) {
                    e.preventDefault();
                    selectSuggestion(suggestions[currentSelection]);
                }
                break;
            case 'Escape':
                hideSuggestions();
                break;
        }
    });
    
    // Hide suggestions when clicking outside the input and suggestions
    document.addEventListener('click', function(e) {
        if (!studentNameInput.contains(e.target) && !studentSuggestions.contains(e.target)) {
            hideSuggestions();
        }
    });
    
    // Fetch student name suggestions
    function fetchSuggestions(query) {
        fetch(`/api/students/suggestions?q=${encodeURIComponent(query)}`)
            .then(response => response.json())
            .then(data => {
                suggestions = data;
                displaySuggestions(data, query);
            })
            .catch(error => {
                console.error('Error fetching suggestions:', error);
            });
    }
    
    // Display suggestions in the dropdown
    function displaySuggestions(data, query) {
        // Clear previous suggestions
        studentSuggestions.innerHTML = '';
        
        if (data.length === 0) {
            hideSuggestions();
            return;
        }
        
        // Create and append suggestion items
        data.forEach((name, index) => {
            const li = document.createElement('li');
            li.className = 'list-group-item suggestion-item';
            li.style.cursor = 'pointer';
            
            // Highlight the matching part of the name
            const lowerName = name.toLowerCase();
            const lowerQuery = query.toLowerCase();
            const startIndex = lowerName.indexOf(lowerQuery);
            
            if (startIndex >= 0) {
                const beforeMatch = name.substring(0, startIndex);
                const match = name.substring(startIndex, startIndex + query.length);
                const afterMatch = name.substring(startIndex + query.length);
                
                li.innerHTML = beforeMatch + '<strong>' + match + '</strong>' + afterMatch;
            } else {
                li.textContent = name;
            }
            
            // Add event listener to select the suggestion
            li.addEventListener('click', () => selectSuggestion(name));
            
            // Add hover effect
            li.addEventListener('mouseenter', () => {
                currentSelection = index;
                highlightSuggestion();
            });
            
            studentSuggestions.appendChild(li);
        });
        
        // Show the suggestions dropdown
        studentSuggestions.style.display = 'block';
    }
    
    // Select a suggestion
    function selectSuggestion(name) {
        studentNameInput.value = name;
        hideSuggestions();
    }
    
    // Hide suggestions dropdown
    function hideSuggestions() {
        studentSuggestions.style.display = 'none';
        currentSelection = -1;
    }
    
    // Highlight the currently selected suggestion
    function highlightSuggestion() {
        const items = studentSuggestions.querySelectorAll('.suggestion-item');
        
        items.forEach((item, index) => {
            if (index === currentSelection) {
                item.classList.add('active', 'bg-primary', 'text-white');
            } else {
                item.classList.remove('active', 'bg-primary', 'text-white');
            }
        });
        
        // Ensure selected item is visible in the dropdown
        if (currentSelection >= 0) {
            const selectedItem = items[currentSelection];
            selectedItem.scrollIntoView({ block: 'nearest' });
        }
    }
    
    // Form submit event handler
    studentSearchForm.addEventListener('submit', function(e) {
        e.preventDefault();
        
        const studentName = studentNameInput.value.trim();
        if (!studentName) return;
        
        hideSuggestions(); // Hide suggestions on submit
        
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