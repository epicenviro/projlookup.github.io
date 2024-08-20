document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const fileButton = document.getElementById('fileButton');
    const fileInfo = document.getElementById('fileInfo');
    const lastUpdated = document.getElementById('lastUpdated');
    const sheetSelect = document.getElementById('sheetSelect');
    const searchInput = document.getElementById('searchInput');
    const searchButton = document.getElementById('searchButton');
    const results = document.getElementById('results');
    const loading = document.getElementById('loading');
    const error = document.getElementById('error');

    let workbook = null;

    // Load cached data if available
    const cachedData = localStorage.getItem('excelData');
    if (cachedData) {
        workbook = XLSX.read(cachedData, { type: 'base64' });
        updateFileInfo(localStorage.getItem('lastUpdated'));
        populateSheetSelect(workbook);
        showSearchElements();
    }

    fileButton.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            loading.style.display = 'block';
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    workbook = XLSX.read(data, {type: 'array'});
                    
                    // Cache the data
                    localStorage.setItem('excelData', XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' }));
                    const now = new Date().toLocaleString();
                    localStorage.setItem('lastUpdated', now);
                    
                    updateFileInfo(now);
                    populateSheetSelect(workbook);
                    showSearchElements();
                    loading.style.display = 'none';
                } catch (error) {
                    console.error("Error reading the file:", error);
                    displayError("Error reading the file. Please make sure it's a valid Excel file.");
                }
            };
            reader.onerror = (error) => {
                console.error("FileReader error:", error);
                displayError("Error reading the file. Please try again.");
            };
            reader.readAsArrayBuffer(file);
        }
    });

    function updateFileInfo(date) {
        lastUpdated.textContent = date;
        fileInfo.style.display = 'block';
        fileButton.textContent = 'Update/New File';
    }

    function showSearchElements() {
        sheetSelect.style.display = 'block';
        searchInput.style.display = 'block';
        searchButton.style.display = 'inline-block';
    }

    function populateSheetSelect(workbook) {
        sheetSelect.innerHTML = '';
        const projectLogSheets = workbook.SheetNames.filter(name => name.startsWith('Project Log '));
        projectLogSheets.forEach(sheetName => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            sheetSelect.appendChild(option);
        });

        if (projectLogSheets.length > 0) {
            sheetSelect.value = projectLogSheets[projectLogSheets.length - 1];
        }
    }

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

        displayResults(searchResults);
    });

    function displayResults(searchResults) {
        // ... (rest of the displayResults function remains the same)
    }

    function displayError(message) {
        error.textContent = message;
        error.style.display = 'block';
        loading.style.display = 'none';
    }

    // Clear file input when the page is refreshed or closed
    window.addEventListener('beforeunload', () => {
        fileInput.value = '';
    });
});