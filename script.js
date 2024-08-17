document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excelFile');
    const sheetSelect = document.getElementById('sheetSelect');
    const searchInput = document.getElementById('searchInput');
    const searchButton = document.getElementById('searchButton');
    const results = document.getElementById('results');

    let workbook = null;

    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        const reader = new FileReader();

        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            workbook = XLSX.read(data, {type: 'array'});
            
            // Populate sheet select
            sheetSelect.innerHTML = '';
            workbook.SheetNames.forEach(sheetName => {
                if (sheetName.startsWith('Project Log ')) {
                    const option = document.createElement('option');
                    option.value = sheetName;
                    option.textContent = sheetName;
                    sheetSelect.appendChild(option);
                }
            });
        };

        reader.readAsArrayBuffer(file);
    });

    searchButton.addEventListener('click', () => {
        if (!workbook) {
            alert('Please select an Excel file first.');
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
                    `).join('')}
                </tbody>
            </table>
        `;

        // Add click event listeners to rows
        document.querySelectorAll('.result-row').forEach(row => {
            row.addEventListener('click', () => {
                const index = row.getAttribute('data-index');
                const fullData = searchResults[index];
                showDetails(fullData);
            });
        });
    });

    function showDetails(rowData) {
        const detailsDiv = document.createElement('div');
        detailsDiv.className = 'details';
        detailsDiv.innerHTML = `
            <h3>Project Number: ${rowData[0]}</h3>
            <p><strong>Description:</strong> ${rowData[2]}</p>
            <p><strong>Client Name:</strong> ${rowData[3]}</p>
            <p><strong>Client ID:</strong> ${rowData[4]}</p>
            <p><strong>Client Contact:</strong> ${rowData[5]}</p>
            <p><strong>Proposal 3:</strong> ${rowData[6]}</p>
            <p><strong>PO #:</strong> ${rowData[7]}</p>
            <p><strong>Amount Charged:</strong> ${rowData[8]}</p>
            <p><strong>Award Date:</strong> ${rowData[10]}</p>
        `;
        
        // Remove any existing details
        const existingDetails = document.querySelector('.details');
        if (existingDetails) {
            existingDetails.remove();
        }
        
        results.appendChild(detailsDiv);
    }
});