// KeyTrack Dashboard JavaScript
document.addEventListener('DOMContentLoaded', function() {
    // Initialize the dashboard
    loadDashboardData();
    
    // Refresh data every 60 seconds
    setInterval(loadDashboardData, 60000);
});

// Maroon color for University of Chicago theme
const MAROON = '#800000';

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
                backgroundColor: ['#28a745', '#17a2b8'],
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
            labels: ['Borrowed', 'Lost', 'Damaged', 'Available'],
            datasets: [{
                label: 'Key Count',
                data: [
                    data.keys_borrowed || 0,
                    data.keys_lost || 0,
                    data.keys_damaged || 0,
                    data.available_keys || 0
                ],
                backgroundColor: [
                    '#ffc107',  // Borrowed - warning
                    '#dc3545',  // Lost - danger
                    '#fd7e14',  // Damaged - orange
                    '#28a745'   // Available - success
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
            renderRoomsTable(data);
        })
        .catch(error => {
            console.error('Error fetching rooms overview:', error);
            document.getElementById('rooms-table').innerHTML = 
                '<tr><td colspan="7" class="text-center text-danger">Failed to load room data. Please try refreshing the page.</td></tr>';
        });
}

// Render the rooms table
function renderRoomsTable(data) {
    const tableBody = document.getElementById('rooms-table');
    
    if (data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="7" class="text-center">No rooms found</td></tr>';
        return;
    }
    
    let html = '';
    data.forEach(room => {
        html += `
            <tr>
                <td>${room.room_id}</td>
                <td>${room.total_keys}</td>
                <td>${room.available_keys}</td>
                <td>${room.collected_keys}</td>
                <td>${room.lost_keys}</td>
                <td>${room.borrowed_keys}</td>
                <td>
                    <a href="/room/${room.room_id}" class="btn btn-sm btn-primary">Details</a>
                </td>
            </tr>
        `;
    });
    
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
        let badgeClass = 'bg-maroon'; // Changed from bg-info to bg-maroon
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