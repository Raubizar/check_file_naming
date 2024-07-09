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
        const analysis = analyzeFileName(result.name);
        row.insertCell(1).textContent = analysis.compliance;
        row.insertCell(2).textContent = analysis.details;
    });
}

function analyzeFileName(fileName) {
    if (!namingConvention) {
        return { compliance: 'No naming convention uploaded', details: 'Please upload a naming convention file' };
    }
    
    let result = 'Ok';
    let details = 'Correct information';

    // Remove file extension
    const dotPosition = fileName.lastIndexOf('.');
    if (dotPosition > 0) {
        fileName = fileName.substring(0, dotPosition);
    }

    // Read parts and delimiter from the naming convention
    const partsCount = parseInt(namingConvention[0][1], 10);
    const delimiter = namingConvention[0][3];

    // Split the file name into parts
    const nameParts = fileName.split(delimiter);

    // Check if delimiter is correct
    if (nameParts.length !== partsCount) {
        result = 'Wrong';
        details = `Delimiter ok; Wrong number of parts`;
        return { compliance: result, details: details };
    }

    // Verify each part against the naming convention
    for (let j = 0; j < nameParts.length; j++) {
        const allowedParts = namingConvention.slice(1, 200).map(row => row[j]);
        let partAllowed = false;

        if (!allowedParts[0]) {
            result = 'Wrong';
            details = `part ${j + 1} "${nameParts[j]}" not found in the naming standard`;
            return { compliance: result, details: details };
        }

        // Specific checks for numeric parts or description
        if (!isNaN(allowedParts[0])) {
            if (nameParts[j].length === parseInt(allowedParts[0], 10)) {
                partAllowed = true;
            }
        } else if (allowedParts[0] === "Description") {
            if (nameParts[j].length >= 3) {
                partAllowed = true;
            }
        } else {
            for (const allowedPart of allowedParts) {
                if (allowedPart === nameParts[j]) {
                    partAllowed = true;
                    break;
                }
            }
        }

        if (!partAllowed) {
            result = 'Wrong';
            details = `part ${j + 1} "${nameParts[j]}" not found in the naming standard`;
            return { compliance: result, details: details };
        }
    }

    return { compliance: result, details: details };
}

function exportResults() {
    const results = [];
    const rows = document.querySelectorAll('#results-table tbody tr');
    results.push({
        name: 'File Name',
        compliance: 'Compliance Status',
        details: 'Details'
    });
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
