document.addEventListener('DOMContentLoaded', () => {
    const sheetSelect = document.getElementById('sheetSelect');
    const searchInput = document.getElementById('searchInput');
    const searchButton = document.getElementById('searchButton');
    const results = document.getElementById('results');
    const loading = document.getElementById('loading');
    const error = document.getElementById('error');

    const GOOGLE_SHEET_ID = '1fqHopVL2NPslUZ-_jKErXsbVO_IntUGG';
    const CORS_PROXY = "https://cors-anywhere.herokuapp.com/";
    const GOOGLE_SHEET_URL = `${CORS_PROXY}https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/export?format=csv`;

    let workbook = null;

    // Load the Google Sheet
    loading.style.display = 'block';
    fetch(GOOGLE_SHEET_URL)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const contentType = response.headers.get("content-type");
            if (contentType && contentType.includes("text/html")) {
                throw new Error("Received HTML instead of CSV. The spreadsheet might not be accessible.");
            }
            return response.text();
        })
        .then(data => {
            workbook = XLSX.read(data, {type: 'string'});
            
            // Populate sheet select
            sheetSelect.innerHTML = '';
            const projectLogSheets = workbook.SheetNames.filter(name => name.startsWith('Project Log '));
            projectLogSheets.forEach(sheetName => {
                const option = document.createElement('option');
                option.value = sheetName;
                option.textContent = sheetName;
                sheetSelect.appendChild(option);
            });

            // Set default to most recent year
            if (projectLogSheets.length > 0) {
                sheetSelect.value = projectLogSheets[projectLogSheets.length - 1];
            }

            loading.style.display = 'none';
        })
        .catch(e => {
            console.error("There was a problem fetching the spreadsheet data:", e);
            error.textContent = "There was a problem loading the data. Please try again later.";
            error.style.display = 'block';
            loading.style.display = 'none';
        });

    searchButton.addEventListener('click', () => {
        if (!workbook) {
            alert('Sheet data is still loading. Please try again in a moment.');
            return;
        }

        const keywords = searchInput.value.toLowerCase().split(' ');
        const sheetName = sheetSelect.value;
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1});

        const searchResults = data.filter((row, index) => {
            if (index < 4 || !row[0]) return false; // Skip header rows and empty project numbers
            const description = (row[2] || '').toString().toLowerCase();
            const clientName = (row[3] || '').toString().toLowerCase();
            return keywords.every(keyword => description.includes(keyword) || clientName.includes(keyword));
        });

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

        // Add click event listeners to rows
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
    });
});