// KeyTrack Dashboard JavaScript
document.addEventListener('DOMContentLoaded', function() {
    // Initialize the dashboard
    loadDashboardData();
    
    // Refresh data every 60 seconds
    setInterval(loadDashboardData, 60000);
    
    // Set up room search functionality
    setupRoomSearch();
});

// UChicago color scheme with transparent maroon variants
const COLORS = {
    maroon: 'rgba(128, 0, 0, 0.8)',       // Phoenix Maroon with transparency
    maroonDark: 'rgba(96, 0, 0, 0.8)',     // Phoenix Maroon Dark with transparency
    maroonLight: 'rgba(154, 0, 0, 0.7)',   // Phoenix Maroon Light with transparency
    greystone: '#D9D9D9',    // Light Greystone
    greystoneDark: '#737373',// Dark Greystone
    footerGrey: '#404040',   // Footer Grey
    goldenrod: '#F3D03E',    // Light Goldenrod
    terracotta: '#ECA154',   // Light Terracotta
    brick: '#B46A55',        // Light Brick
    ivy: '#A9C47F',          // Light Ivy
    forest: '#9CAF88',       // Light Forest
    lake: 'rgba(62, 177, 200, 0.8)',         // Light Lake with transparency
    violet: '#86647A'        // Light Violet
};

// Store all rooms data globally so we can filter without re-fetching
let allRoomsData = [];

// Main function to load all dashboard data
function loadDashboardData() {
    fetchKeyStatistics();
    fetchRoomsOverview();
    fetchRecentActivity();
}

// Fetch and display key statistics
function fetchKeyStatistics() {
    fetch('/api/keys/statistics')
        .then(response => response.json())
        .then(data => {
            updateStatisticsOverview(data);
            renderAvailableKeysChart(data);
            renderKeyStatusChart(data);
        })
        .catch(error => {
            console.error('Error fetching key statistics:', error);
            document.getElementById('stats-overview').innerHTML = 
                '<div class="alert alert-danger">Failed to load key statistics. Please try refreshing the page.</div>';
        });
}

// Update statistics overview
function updateStatisticsOverview(data) {
    const statsElement = document.getElementById('stats-overview');
    statsElement.innerHTML = `
        <p><strong>Total Keys:</strong> ${data.total_keys}</p>
        <p><strong>Available Keys:</strong> ${data.available_keys} (${Math.round((data.available_keys / data.total_keys) * 100)}%)</p>
        <p><strong>Keys Currently Out:</strong> ${data.keys_out}</p>
    `;
}

// Render the Available Keys chart
function renderAvailableKeysChart(data) {
    const ctx = document.getElementById('availableKeysChart').getContext('2d');
    
    // Destroy existing chart if it exists
    if (window.availableKeysChart instanceof Chart) {
        window.availableKeysChart.destroy();
    }
    
    window.availableKeysChart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: ['Available', 'In Use'],
            datasets: [{
                data: [data.available_keys, data.keys_out],
                backgroundColor: [COLORS.ivy, COLORS.lake],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom'
                }
            }
        }
    });
}

// Render the Key Status chart
function renderKeyStatusChart(data) {
    const ctx = document.getElementById('keyStatusChart').getContext('2d');
    
    // Destroy existing chart if it exists
    if (window.keyStatusChart instanceof Chart) {
        window.keyStatusChart.destroy();
    }
    
    window.keyStatusChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Borrowed', 'Lost', 'Available'],
            datasets: [{
                label: 'Key Count',
                data: [
                    data.keys_borrowed || 0,
                    data.keys_lost || 0,
                    data.available_keys || 0
                ],
                backgroundColor: [
                    COLORS.goldenrod,  // Borrowed - light goldenrod
                    COLORS.maroonLight,  // Lost - phoenix maroon light with transparency
                    COLORS.ivy   // Available - light ivy
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true
                }
            },
            plugins: {
                legend: {
                    display: false
                }
            }
        }
    });
}

// Fetch and display rooms overview
function fetchRoomsOverview() {
    fetch('/api/rooms/overview')
        .then(response => response.json())
        .then(data => {
            // Store the complete rooms data for filtering
            allRoomsData = data;
            
            // Get the current search term
            const searchTerm = document.getElementById('roomSearchInput').value.trim().toLowerCase();
            
            // Render the table with appropriate filtering
            renderRoomsTable(data, searchTerm);
        })
        .catch(error => {
            console.error('Error fetching rooms overview:', error);
            document.getElementById('rooms-table').innerHTML = 
                '<tr><td colspan="7" class="text-center text-danger">Failed to load room data. Please try refreshing the page.</td></tr>';
        });
}

// Render the rooms table
function renderRoomsTable(data, searchTerm = '') {
    const tableBody = document.getElementById('rooms-table');
    
    if (data.length === 0) {
        tableBody.innerHTML = searchTerm 
            ? '<tr><td colspan="7" class="text-center">No rooms matching "' + searchTerm + '" found</td></tr>'
            : '<tr><td colspan="7" class="text-center">No rooms found</td></tr>';
        return;
    }
    
    let html = '';
    let roomsToDisplay = data;
    
    if (searchTerm) {
        // If searching, filter the rooms
        roomsToDisplay = data.filter(room => {
            const roomId = String(room.room_id).toLowerCase();
            return roomId.includes(searchTerm);
        });
        
        // Check if any rooms match the search
        if (roomsToDisplay.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="7" class="text-center">No rooms matching "' + searchTerm + '" found</td></tr>';
            return;
        }
    } else {
        // If not searching, limit to 10 rooms
        roomsToDisplay = data.slice(0, 10);
    }
    
    // Generate the HTML for each room row
    roomsToDisplay.forEach(room => {
        // Properly encode the room_id for the URL to prevent issues with spaces or special characters
        const encodedRoomId = encodeURIComponent(room.room_id);
        
        html += `
            <tr>
                <td>${room.room_id}</td>
                <td>${room.total_keys}</td>
                <td>${room.available_keys}</td>
                <td>${room.collected_keys}</td>
                <td>${room.lost_keys}</td>
                <td>${room.borrowed_keys}</td>
                <td>
                    <a href="/room/${encodedRoomId}" class="btn btn-sm btn-primary">Details</a>
                </td>
            </tr>
        `;
    });
    
    // Add a summary row if not in search mode and there are more than 10 rooms
    if (!searchTerm && data.length > 10) {
        html += `
            <tr>
                <td colspan="7" class="text-center text-muted">
                    Showing 10 of ${data.length} rooms. Use the search to find specific rooms.
                </td>
            </tr>
        `;
    } else if (searchTerm) {
        // If in search mode, show how many results were found
        html += `
            <tr>
                <td colspan="7" class="text-center text-muted">
                    Found ${roomsToDisplay.length} room(s) matching "${searchTerm}".
                </td>
            </tr>
        `;
    }
    
    tableBody.innerHTML = html;
}

// Fetch and display recent activity
function fetchRecentActivity() {
    fetch('/api/activity/recent')
        .then(response => response.json())
        .then(data => {
            renderRecentActivity(data);
        })
        .catch(error => {
            console.error('Error fetching recent activity:', error);
            document.getElementById('recent-activity').innerHTML = 
                '<li class="list-group-item text-danger">Failed to load recent activity. Please try refreshing the page.</li>';
        });
}

// Render recent activity
function renderRecentActivity(data) {
    const activityList = document.getElementById('recent-activity');
    
    if (data.length === 0) {
        activityList.innerHTML = '<li class="list-group-item text-center">No recent activity</li>';
        return;
    }
    
    let html = '';
    data.forEach(activity => {
        let badgeClass = 'bg-maroon';
        if (activity.action === 'borrowed') badgeClass = 'bg-warning';
        if (activity.action === 'returned') badgeClass = 'bg-success';
        if (activity.action === 'lost') badgeClass = 'bg-danger';
        if (activity.action === 'collected') badgeClass = 'bg-maroon';
        
        html += `
            <li class="list-group-item d-flex justify-content-between align-items-center">
                ${activity.description}
                <div>
                    <span class="badge ${badgeClass}">${activity.action}</span>
                    <small class="text-muted ms-2">${new Date(activity.timestamp).toLocaleString()}</small>
                </div>
            </li>
        `;
    });
    
    activityList.innerHTML = html;
}

// Set up the search functionality
function setupRoomSearch() {
    const searchInput = document.getElementById('roomSearchInput');
    const clearButton = document.getElementById('clearSearchButton');
    
    searchInput.addEventListener('input', function() {
        const searchTerm = this.value.trim().toLowerCase();
        filterRooms(searchTerm);
    });
    
    clearButton.addEventListener('click', function() {
        searchInput.value = '';
        filterRooms('');
    });
}

// Filter rooms based on search term
function filterRooms(searchTerm) {
    renderRoomsTable(allRoomsData, searchTerm);
}