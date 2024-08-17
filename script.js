const SPREADSHEET_ID = '1fqHopVL2NPslUZ-_jKErXsbVO_IntUGG';
const API_KEY = 'AIzaSyD1QnyX9WpD8Cj2lSGctK9UKIuqDNsBGws'; // Replace with your actual API key

document.addEventListener('DOMContentLoaded', () => {
    const runButton = document.getElementById('runButton');
    const sheetSelect = document.getElementById('sheetSelect');
    const searchInput = document.getElementById('searchInput');
    const searchButton = document.getElementById('searchButton');
    const results = document.getElementById('results');
    const loading = document.getElementById('loading');
    const error = document.getElementById('error');

    let sheets = [];

    runButton.addEventListener('click', initializeApp);

    function initializeApp() {
        loading.style.display = 'block';
        gapi.load('client', initializeGapiClient);
    }

    async function initializeGapiClient() {
        await gapi.client.init({
            apiKey: API_KEY,
            discoveryDocs: ["https://sheets.googleapis.com/$discovery/rest?version=v4"],
        });
        loadSheets();
    }

    async function loadSheets() {
        try {
            const response = await gapi.client.sheets.spreadsheets.get({
                spreadsheetId: SPREADSHEET_ID
            });
            sheets = response.result.sheets;
            populateSheetSelect(sheets);
            loading.style.display = 'none';
            sheetSelect.style.display = 'block';
            searchInput.style.display = 'block';
            searchButton.style.display = 'inline-block';
            runButton.style.display = 'none';
        } catch (err) {
            console.error("Error loading sheets", err);
            displayError("Error loading spreadsheet data. Please try again later.");
        }
    }

    function populateSheetSelect(sheets) {
        sheetSelect.innerHTML = '';
        sheets.forEach(sheet => {
            if (sheet.properties.title.startsWith('Project Log ')) {
                const option = document.createElement('option');
                option.value = sheet.properties.title;
                option.textContent = sheet.properties.title;
                sheetSelect.appendChild(option);
            }
        });
    }

    searchButton.addEventListener('click', handleSearch);

    async function handleSearch() {
        const keywords = searchInput.value.toLowerCase().split(' ');
        const sheetName = sheetSelect.value;
        
        loading.style.display = 'block';
        results.innerHTML = '';
        error.style.display = 'none';

        try {
            const response = await gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: SPREADSHEET_ID,
                range: `${sheetName}!A:K`
            });
            const data = response.result.values;
            const searchResults = data.filter((row, index) => {
                if (index < 4 || !row[0]) return false; // Skip header rows and empty project numbers
                const description = (row[2] || '').toString().toLowerCase();
                const clientName = (row[3] || '').toString().toLowerCase();
                return keywords.every(keyword => description.includes(keyword) || clientName.includes(keyword));
            });
            displayResults(searchResults);
            loading.style.display = 'none';
        } catch (err) {
            console.error("Error fetching data", err);
            displayError("Error fetching data. Please try again later.");
        }
    }

    function displayResults(searchResults) {
        results.innerHTML = `
            <table>
                <thead>
                    <tr>
                        <th>Project Number</th>
                        <th>Description</th>
                        <th>Client Name</th>
                    </tr>
                </thead>
                <tbody>
                    ${searchResults.map((row, index) => `
                        <tr class="result-row" data-index="${index}">
                            <td>${row[0]}</td>
                            <td>${row[2]}</td>
                            <td>${row[3]}</td>
                        </tr>
                        <tr class="details-row" id="details-${index}" style="display: none;">
                            <td colspan="3">
                                <div class="details">
                                    <p><strong>Client ID:</strong> ${row[4]}</p>
                                    <p><strong>Client Contact:</strong> ${row[5]}</p>
                                    <p><strong>Proposal 3:</strong> ${row[6]}</p>
                                    <p><strong>PO #:</strong> ${row[7]}</p>
                                    <p><strong>Amount Charged:</strong> ${row[8]}</p>
                                    <p><strong>Award Date:</strong> ${row[10]}</p>
                                </div>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        `;

        document.querySelectorAll('.result-row').forEach(row => {
            row.addEventListener('click', () => {
                const index = row.getAttribute('data-index');
                const detailsRow = document.getElementById(`details-${index}`);
                if (detailsRow.style.display === 'none') {
                    detailsRow.style.display = 'table-row';
                } else {
                    detailsRow.style.display = 'none';
                }
            });
        });
    }

    function displayError(message) {
        error.textContent = message;
        error.style.display = 'block';
        loading.style.display = 'none';
    }
});