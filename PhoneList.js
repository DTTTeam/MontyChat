let officeInitialized = false;

function initializeOffice() {
    if (!officeInitialized) {
        Office.onReady(function (info) {
            $(document).ready(function () {
                initPhoneNumberDropdown();
            });
        });
        officeInitialized = true;
    }
}
initializeOffice();

function initPhoneNumberDropdown() {
    const phoneData = JSON.parse(sessionStorage.getItem('phoneData'));
    const phoneDropdown = document.getElementById('phoneNumberDropdown');

    if (!phoneData || phoneData.length === 0) {
        showNotification("No phone numbers available.", "error");
        return;
    }

    phoneData.forEach(item => {
        let option = document.createElement('option');
        option.text = item.display_phone_number;
        option.value = item.waba_id;
        phoneDropdown.appendChild(option);
    });

    phoneDropdown.addEventListener('change', function () {
        hideStatisticsAndCost();
        const selectedPhoneNumberId = this.value;
        const secretKey = localStorage.getItem('authToken');
        if (selectedPhoneNumberId) {
            fetchData(selectedPhoneNumberId, secretKey);
        }
    });

    phoneDropdown.dispatchEvent(new Event('change'));
}

function populateTemplateDropdown(headers, templates) {
    const templateDropdown = document.getElementById('templateDropdown');
    templateDropdown.innerHTML = '';

    if (templates.length === 0) {
        showNotification("No templates available.", "error");
        return;
    }

    templates.forEach(template => {
        let option = document.createElement('option');
        option.value = template.id;
        option.text = template.name;
        templateDropdown.appendChild(option);
    });

    templateDropdown.addEventListener('change', function () {
        const selectedTemplateId = this.value;
        const selectedTemplate = templates.find(t => t.id === selectedTemplateId);
        if (selectedTemplate) {
            clearStatusColumn();
            clearStatisticsAndGraph();
            updateExcelColumns(selectedTemplate);
            createDropdownsFromComponents(selectedTemplate.components, headers);
            populateVariableDropdowns();
            hideStatisticsAndCost();
        }
    });

    templateDropdown.dispatchEvent(new Event('change'));
}

function fetchData(selectedPhoneNumberId, secretKey) {
    showLoadingSpinner();

    const apiUrl = `https://prod-150.westeurope.logic.azure.com:443/workflows/eb6e5ff3a3be4613a4b2cda11b20412e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=z-z0NWX7WWJ_hFkyVVRbmFPvuqjmGVUQACh4znGqV6Q`;
    const bodyData = {
        SecretKey: secretKey,
        wabaid: selectedPhoneNumberId
    };

    fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(bodyData)
    })
        .then(response => {
            if (!response.ok) {
                throw new Error('Invalid response from server');
            }
            return response.json();
        })
        .then(responseData => {
            hideLoadingSpinner();
            if (responseData && Array.isArray(responseData.data)) {
                populateTemplateDropdown(responseData.headers, responseData.data);
            } else {
                showNotification("No templates found for the selected phone number.", "error");
            }
        })
        .catch(error => {
            hideLoadingSpinner();
            showNotification("Failed to fetch templates: " + error.message, "error");
        });
}

function updateExcelColumns(selectedTemplate) {
    Excel.run(function (context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange("A1:Z1");
        range.clear();

        const phoneNumberColumn = sheet.getRange("A1");
        phoneNumberColumn.values = [["Phone Number"]];

        const statusColumn = sheet.getRange("B1");
        statusColumn.values = [["Status"]];
        statusColumn.format.autofitColumns();

        if (selectedTemplate && selectedTemplate.components) {
            let columnIndex = 2;

            selectedTemplate.components.forEach(component => {
                if (component.type === 'HEADER' && component.text.includes('{{1}}')) {
                    const headerColumn = sheet.getRange(String.fromCharCode(65 + columnIndex) + "1");
                    headerColumn.values = [["Header"]];
                    headerColumn.format.autofitColumns();
                    columnIndex++;
                }

                if (component.type === 'BODY') {
                    const matches = component.text.match(/{{\d+}}/g);
                    if (matches) {
                        matches.forEach((match, index) => {
                            const bodyColumn = sheet.getRange(String.fromCharCode(65 + columnIndex) + "1");
                            bodyColumn.values = [[`Body ${index + 1}`]];
                            bodyColumn.format.autofitColumns();
                            columnIndex++;
                        });
                    }
                }
            });
        }

        return context.sync();
    }).catch(function (error) {
        console.error('Error updating Excel columns:', error);
    });
}

function populateVariableDropdowns() {
    Excel.run(function (context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange("A1:Z1");
        range.load("values");

        return context.sync().then(function () {
            const headers = range.values[0];
            updateVariableDropdowns(headers);
        });
    }).catch(function (error) {
        console.error('Error populating variable dropdowns:', error);
    });
}

function updateVariableDropdowns(headers) {
    const dropdowns = document.querySelectorAll('.variable-dropdown');
    dropdowns.forEach(dropdown => {
        dropdown.innerHTML = '';
        headers.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            option.textContent = header;
            dropdown.appendChild(option);
        });
    });
}

function createDropdownsFromComponents(components, headers) {
    const container = document.getElementById('componentContainer');
    if (!container) return;
    container.innerHTML = '';

    components.forEach((component, index) => {
        const section = document.createElement('div');
        section.className = 'template-section';
        container.appendChild(section);

        const title = document.createElement('h3');
        title.textContent = component.type;
        section.appendChild(title);

        const textPreview = document.createElement('p');
        const text = component.text || '';
        textPreview.innerHTML = text.replace(/\{\{\d+\}\}/g, '<span class="placeholder-highlight">$&</span>');
        section.appendChild(textPreview);
    });
}

function showbar() {
    const progressBar = document.getElementById('progressBar');
    const progressBarFill = document.getElementById('progressBarFill');
    const progressBarText = document.getElementById('progressBarText');
    progressBar.style.display = 'block';
    progressBarFill.style.width = '0%';
    progressBarText.textContent = '0%';
}

async function sendMessages() {
    showLoadingSpinner();
    document.querySelector('.tab-container').style.display = 'none';

    Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const headerRange = sheet.getRange("A1:Z1");
        headerRange.load("values");

        const dataRange = sheet.getRange("A2:Z100");
        dataRange.load("values");

        await context.sync();

        const headers = headerRange.values[0];
        const data = dataRange.values;
        const phoneNumberIndex = headers.indexOf("Phone Number");
        const statusIndex = headers.indexOf("Status");

        if (phoneNumberIndex === -1) {
            console.error("No 'Phone Number' column found in headers.");
            hideSpinner();
            return;
        }

        if (statusIndex === -1) {
            console.error("No 'Status' column found in headers.");
            hideSpinner();
            return;
        }

        const nonEmptyPhoneNumbers = data.filter(row => row[phoneNumberIndex]);
        if (nonEmptyPhoneNumbers.length === 0) {
            showNotification("No phone numbers available.", "error");
            hideSpinner();
            return;
        }

        let hasError = false;
        nonEmptyPhoneNumbers.forEach((row, rowIndex) => {
            headers.forEach((header, index) => {
                if ((header.includes("Header") || header.includes("Body")) && !row[index]) {
                    showNotification(`Missing ${header} value in row ${rowIndex + 2}`, "error");
                    hasError = true;
                }
            });
        });

        if (hasError) {
            hideSpinner();
            return;
        }

        showbar();
        const templateDropdown = document.getElementById('templateDropdown');
        const templateName = templateDropdown.options[templateDropdown.selectedIndex].text;

        let totalNumbers = nonEmptyPhoneNumbers.length;
        let sentCount = 0;
        let failedCount = 0;

        updateStatistics(totalNumbers, sentCount, failedCount);

        for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
            const row = data[rowIndex];
            const phoneNumber = row[phoneNumberIndex];
            if (!phoneNumber) {
                continue;
            }

            const components = [];
            let currentBodyComponent = {
                type: "body",
                parameters: []
            };

            headers.forEach((header, index) => {
                const value = row[index];
                if (header.includes("Header")) {
                    components.push({
                        type: "header",
                        parameters: [{
                            type: "text",
                            text: value.toString()
                        }]
                    });
                } else if (header.includes("Body")) {
                    currentBodyComponent.parameters.push({
                        type: "text",
                        text: value ? value.toString().trim() : ''
                    });
                }
            });

            if (currentBodyComponent.parameters.length > 0) {
                components.push(currentBodyComponent);
            }

            const requestBody = {
                "messaging_product": "whatsapp",
                "recipient_type": "individual",
                "to": phoneNumber,
                "type": "template",
                "template": {
                    "name": templateName,
                    "language": {
                        "code": "en_US"
                    },
                    "components": components
                }
            };

            console.log("Request Body:", JSON.stringify(requestBody, null, 2));

            try {
                const response = await fetch('https://graph.facebook.com/v18.0/165047336701866/messages', {
                    method: 'POST',
                    headers: {
                        'Authorization': 'Bearer EAAS4mS9ezTwBOzVgqqQe9eqtsdlppZBpcnZBSpuBYZA8OZCRCenrNOZAQinhSd3uwvXfTVpiIiZAZBtNt1SZB6FxbzY9ixwm8pgyOPMHfXXJW19QKZAwteUVpRERbeAQAWLdYVY9ILXkLZC8ahUlKMz1ZBsF8GJmk7qgBDRcpeYS1ha3C1nhVHxXTc5zDKX7ekAXdvWyZB18t1YIBYGgAalQXFutwt0sRkEGvCUsxbghkjvaLuVgplAvXIaY',
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(requestBody)
                });

                const responseData = await response.json();
                console.log('Response Data:', responseData);

                if (responseData.messages && responseData.messages.length > 0) {
                    data[rowIndex][statusIndex] = `Sent`;
                    sentCount++;
                } else {
                    data[rowIndex][statusIndex] = 'Failed';
                    failedCount++;
                }

            } catch (error) {
                console.error('Failed to send API request:', error);
                data[rowIndex][statusIndex] = 'Failed';
                failedCount++;
            }

            updateStatistics(totalNumbers, sentCount, failedCount);

            const rangeToUpdate = sheet.getRange(`A${rowIndex + 2}:Z${rowIndex + 2}`);
            rangeToUpdate.values = [data[rowIndex]];
            await context.sync();

            // Update progress bar
            const progressPercent = ((rowIndex + 1) / totalNumbers) * 100;
            progressBarFill.style.width = `${progressPercent}%`;
            progressBarText.textContent = `${Math.round(progressPercent)}%`;
        }

        hideSpinner();
        document.querySelector('.tab-container').style.display = 'block';

    }).catch((error) => {
        console.error('Error accessing Excel file:', error);
        hideSpinner();
    });
}

function hideSpinner() {
    document.getElementById('loadingSpinner').style.display = 'none';
    document.getElementById('loadingMessage').style.display = 'none';
}

function showLoadingSpinner() {
    document.getElementById('loadingSpinner').style.display = 'block';
    document.getElementById('loadingMessage').style.display = 'block';
}

function clearStatusColumn() {
    Excel.run(function (context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const statusRange = sheet.getRange("B2:B100");
        statusRange.clear();

        return context.sync().then(function () {
            console.log("Status column cleared successfully.");
        });
    }).catch(function (error) {
        console.error('Error clearing status column:', error);
    });
}

function updateStatistics(totalNumbers, sentCount, failedCount) {
    document.getElementById('totalNumbers').textContent = `Total Phone Numbers: ${totalNumbers}`;
    document.getElementById('sentNumbers').textContent = `Sent: ${sentCount}`;
    document.getElementById('failedNumbers').textContent = `Failed: ${failedCount}`;
    updateTotalCost(sentCount);
    updateStatisticsGraph(totalNumbers, sentCount, failedCount);
}

function updateTotalCost(sentCount) {
    const costPerMessage = parseFloat(document.getElementById('costPerMessage').value) || 0;
    const totalCost = (sentCount * costPerMessage).toFixed(2);
    document.getElementById('totalCost').textContent = `Total Cost: $${totalCost}`;
}

function clearStatistics() {
    document.getElementById('totalNumbers').textContent = `Total Phone Numbers: 0`;
    document.getElementById('sentNumbers').textContent = `Sent: 0`;
    document.getElementById('failedNumbers').textContent = `Failed: 0`;
    document.getElementById('totalCost').textContent = `Total Cost: $0.00`;
}

function showConfirmationModal(callback) {
    document.getElementById('confirmationModal').style.display = "block";
    document.getElementById('confirmYes').onclick = function () {
        callback(true);
        hideConfirmationModal();
    };
    document.getElementById('confirmNo').onclick = function () {
        callback(false);
        hideConfirmationModal();
    };
}

function hideConfirmationModal() {
    document.getElementById('confirmationModal').style.display = "none";
}

function updateStatisticsGraph(total, sent, failed) {
    const ctx = document.getElementById('statisticsGraph').getContext('2d');
    if (window.myChart) {
        window.myChart.destroy();
    }
    window.myChart = new Chart(ctx, {
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

function clearExcelData() {
    Excel.run(function (context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange("B2:Z100");
        range.clear();
        const statusColumn = sheet.getRange("B2:B100");
        statusColumn.clear();

        return context.sync();
    }).then(function () {
        console.log('Excel data and status column cleared.');
    }).catch(function (error) {
        console.error('Error clearing Excel data:', error);
    });
}

function clearStatisticsAndGraph() {
    document.getElementById('totalNumbers').textContent = 0;
    document.getElementById('sentNumbers').textContent = 0;
    document.getElementById('failedNumbers').textContent = 0;
    document.getElementById('totalCost').textContent = '$0.00';
    updateStatisticsGraph(0, 0, 0);
}

function checkExcelData(callback) {
    Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const dataRange = sheet.getRange("B2:Z100");
        dataRange.load("values");

        await context.sync();

        const hasData = dataRange.values.some(row => row.some(cell => cell !== null && cell !== ""));
        callback(hasData);
    }).catch((error) => {
        console.error('Error checking Excel data:', error);
    });
}

document.getElementById('phoneNumberDropdown').addEventListener('change', function () {
    checkExcelData((hasData) => {
        if (hasData) {
            showConfirmationModal((confirm) => {
                if (confirm) {
                    clearExcelData();
                    clearStatistics();
                }
                clearStatisticsAndGraph();
            });
        } else {
            clearStatisticsAndGraph();
        }
    });
});

document.getElementById('templateDropdown').addEventListener('change', function () {
    checkExcelData((hasData) => {
        if (hasData) {
            showConfirmationModal((confirm) => {
                if (confirm) {
                    clearExcelData();
                    clearStatistics();
                }
                clearStatisticsAndGraph();
            });
        } else {
            clearStatisticsAndGraph();
        }
    });
});

document.getElementById('sendBtn').addEventListener('click', sendMessages);

document.querySelectorAll('.tab-button').forEach(button => {
    button.addEventListener('click', function (event) {
        openTab(event, button.getAttribute('data-tab'));
    });
});

function openTab(evt, tabName) {
    var i, tabcontent, tablinks;
    tabcontent = document.getElementsByClassName("tab-content");
    for (i = 0; i < tabcontent.length; i++) {
        tabcontent[i].style.display = "none";
    }
    tablinks = document.getElementsByClassName("tab-button");
    for (i = 0; i < tablinks.length; i++) {
        tablinks[i].className = tablinks[i].className.replace(" active", "");
    }
    document.getElementById(tabName).style.display = "block";
    evt.currentTarget.className += " active";

    if (tabName === 'Statistics') {
        const totalNumbers = (document.getElementById('totalNumbers')?.textContent.split(':')[1] || "0").trim();
        const sentNumbers = (document.getElementById('sentNumbers')?.textContent.split(':')[1] || "0").trim();
        const failedNumbers = (document.getElementById('failedNumbers')?.textContent.split(':')[1] || "0").trim();
        const totalCost = (document.getElementById('totalCost')?.textContent.split('$')[1] || "0.00").trim();

        console.log(`Navigating to statistics.html with data: total=${totalNumbers}, sent=${sentNumbers}, failed=${failedNumbers}, cost=${totalCost}`);

        window.location.href = `https://dttteam.github.io/MontyChat//statistics.html?total=${totalNumbers}&sent=${sentNumbers}&failed=${failedNumbers}&cost=${totalCost}`;
    }
}

function hideLoadingSpinner() {
    document.getElementById('loadingSpinner').style.display = 'none';
    document.getElementById('loadingMessage').style.display = 'none';
}

function hideStatisticsAndCost() {
    const tabContainer = document.querySelector('.tab-container');
    const progressBar = document.getElementById('progressBar');

    if (tabContainer && tabContainer.style.display !== 'none') {
        tabContainer.style.display = 'none';
    }
    if (progressBar && progressBar.style.display !== 'none') {
        progressBar.style.display = 'none';
    }
}

function showNotification(message, type) {
    const notification = document.getElementById('notification');
    notification.textContent = message;
    notification.className = 'notification';
    notification.classList.add(type === "error" ? 'error' : 'success');
    notification.classList.add('show');
    setTimeout(() => {
        notification.classList.remove('show');
    }, 3000);
}
