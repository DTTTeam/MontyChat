Office.onReady((info) => {
    // Office is ready
    $(document).ready(function () {
        $('#login-button').click(function () {
            var secretKey = $('#secret-key').val();
            if (!secretKey) {
                showNotification('Secret key is required!', "error");
            } else {
                callApi(secretKey);
            }
        });
    });
});

function callApi(secretKey) {
    $('#loadingSpinner').show();
    const requestOptions = {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ "SecretKey": secretKey }),
        redirect: "follow"
    };

    fetch("https://prod-62.westeurope.logic.azure.com:443/workflows/3224497220d845a587499c61c393b7f7/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=iHuKcW26CtwasDukBkYOfvx7X1H6oVIV99cvtfzDUu8", requestOptions)
        .then(response => {
            if (!response.ok) {
                throw new Error('Invalid response from server');
            }
            return response.json();
        })
        .then(data => {
            if (data && data.length > 0) {
                localStorage.setItem('authToken', secretKey);
                sessionStorage.setItem('phoneData', JSON.stringify(data));
                window.location.href = "https://dttteam.github.io/MontyChatWeb/Phonelist.html";
            } else {
                showNotification('No valid data found.', "error");
            }
        })
        .catch(error => {
            showNotification(error.message, "error");
        })
        .finally(() => {
            $('#loadingSpinner').hide(); // Hide the spinner
        });
}

function showNotification(message, type) {
    const notification = document.getElementById('notification');
    if (type === "error") {
        notification.style.backgroundColor = "red";
    } else if (type === "success") {
        notification.style.backgroundColor = "green";
    }
    notification.textContent = message;
    notification.classList.add('show');
    setTimeout(() => {
        notification.classList.remove('show');
    }, 3000);
}
