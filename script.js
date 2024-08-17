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

        results.innerHTML = searchResults.map(row => `
            <div class="result-item">
                <h3>Project Number: ${row[0]}</h3>
                <p><strong>Description:</strong> ${row[2]}</p>
                <p><strong>Client Name:</strong> ${row[3]}</p>
                <p><strong>Client ID:</strong> ${row[4]}</p>
                <p><strong>Client Contact:</strong> ${row[5]}</p>
                <p><strong>Proposal 3:</strong> ${row[6]}</p>
                <p><strong>PO #:</strong> ${row[7]}</p>
                <p><strong>Amount Charged:</strong> ${row[8]}</p>
                <p><strong>Award Date:</strong> ${row[10]}</p>
            </div>
        `).join('');
    });
});