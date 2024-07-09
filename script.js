document.getElementById('folder-select-button').addEventListener('click', async () => {
    const directoryHandle = await window.showDirectoryPicker();
    const results = [];
    for await (const entry of directoryHandle.values()) {
        if (entry.kind === 'file') {
            const file = await entry.getFile();
            results.push({ name: file.name, compliance: 'Pending', details: 'Pending' });
        }
    }
    displayResults(results);
});

document.getElementById('file-input').addEventListener('change', handleFileUpload);
document.getElementById('export-button').addEventListener('click', exportResults);

let namingConvention = null;

async function handleFileUpload(event) {
    const file = event.target.files[0];
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    namingConvention = XLSX.utils.sheet_to_json(sheet, { header: 1 });
}

function displayResults(results) {
    const tbody = document.getElementById('results-table').querySelector('tbody');
    tbody.innerHTML = '';
    results.forEach(result => {
        const row = tbody.insertRow();
        row.insertCell(0).textContent = result.name;
        row.insertCell(1).textContent = result.compliance;
        row.insertCell(2).textContent = result.details;
    });
}

function exportResults() {
    const results = [];
    const rows = document.querySelectorAll('#results-table tbody tr');
    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        results.push({
            name: cells[0].textContent,
            compliance: cells[1].textContent,
            details: cells[2].textContent
        });
    });
    const csvContent = "data:text/csv;charset=utf-8,"
        + results.map(e => `${e.name},${e.compliance},${e.details}`).join("\n");

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "results.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
