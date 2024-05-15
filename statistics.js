// Function to navigate back to the phone list
function goBack() {
    window.location.href = 'https://dttteam.github.io/MontyChat/PhoneList.html';
}

// Function to get URL parameters
function getUrlParams() {
    const params = {};
    const queryString = window.location.search.substring(1);
    const regex = /([^&=]+)=([^&]*)/g;
    let match;

    while (match = regex.exec(queryString)) {
        params[decodeURIComponent(match[1])] = decodeURIComponent(match[2]);
    }

    return params;
}


const params = getUrlParams();
console.log('Received parameters:', params); // Debug log for received parameters

updateStatistics(
    parseInt(params.total) || 0,
    parseInt(params.sent) || 0,
    parseInt(params.failed) || 0,
    parseFloat(params.cost) || 0.0
);


function updateStatistics(total, sent, failed, totalCost) {
    document.getElementById('totalNumbers').textContent = `Total Phone Numbers: ${total}`;
    document.getElementById('sentNumbers').textContent = `Sent: ${sent}`;
    document.getElementById('failedNumbers').textContent = `Failed: ${failed}`;
    document.getElementById('totalCost').textContent = `Total Cost: $${totalCost.toFixed(2)}`;

    const ctx = document.getElementById('statisticsGraph').getContext('2d');
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Total', 'Sent', 'Failed'],
            datasets: [{
                label: 'Messages',
                data: [total, sent, failed],
                backgroundColor: ['#ffcd56', '#4caf50', '#f44336'],
                borderColor: ['#ffcd56', '#4caf50', '#f44336'],
                borderWidth: 1
            }]
        },
        options: {
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}
