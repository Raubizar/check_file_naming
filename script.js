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
    if (!file) return;

    console.log('Reading file:', file.name);
    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        namingConvention = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        console.log('Naming convention loaded:', namingConvention);
    } catch (error) {
        console.error('Error reading file:', error);
    }
}

function displayResults(results) {
    const tbody = document.getElementById('results-table').querySelector('tbody');
    tbody.innerHTML = '';

    // Remove and recreate thead to ensure headers are correctly styled
    const thead = document.getElementById('results-table').querySelector('thead');
    thead.innerHTML = '';
    const headerRow = thead.insertRow();
    headerRow.insertCell(0).textContent = 'File Name';
    headerRow.insertCell(1).textContent = 'Compliance Status';
    headerRow.insertCell(2).textContent = 'Details';

    // Add the header class to each cell
    headerRow.cells[0].classList.add('header-cell');
    headerRow.cells[1].classList.add('header-cell');
    headerRow.cells[2].classList.add('header-cell');

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
    
    let delimiterCompliance = 'Ok';
    let partsCountCompliance = 'Ok';
    let partsCompliance = 'Ok';
    let details = '';

    // Remove file extension
    const dotPosition = fileName.lastIndexOf('.');
    if (dotPosition > 0) {
        fileName = fileName.substring(0, dotPosition);
    }

    // Read parts and delimiter from the naming convention
    const partsCount = parseInt(namingConvention[0][1], 10);
    const delimiter = namingConvention[0][3];

    // Split the file name into parts using the specified delimiter
    const nameParts = fileName.split(delimiter);

    // Check if the delimiter is correct
    const expectedDelimiters = partsCount - 1;
    const actualDelimiters = (fileName.match(new RegExp(`\\${delimiter}`, 'g')) || []).length;
    if (actualDelimiters === expectedDelimiters) {
        details += 'Delimiter correct; ';
    } else {
        delimiterCompliance = 'Wrong';
        details += 'Delimiter wrong; ';
    }

    // Check if the number of parts is correct
    if (nameParts.length === partsCount) {
        details += 'Number of parts correct; ';
    } else {
        partsCountCompliance = 'Wrong';
        details += `Number of parts wrong (${nameParts.length}); `;
    }

    // Verify each part against the naming convention
    let nonCompliantParts = [];
    for (let j = 0; j < nameParts.length; j++) {
        const allowedParts = namingConvention.slice(1).map(row => row[j]);
        let partAllowed = false;

        if (!allowedParts[0]) {
            nonCompliantParts.push(nameParts[j]);
            continue;
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
            nonCompliantParts.push(nameParts[j]);
        }
    }

    if (nonCompliantParts.length > 0) {
        partsCompliance = 'Wrong';
        details += `Parts not compliant: ${nonCompliantParts.join(', ')}`;
    }

    let compliance = 'Ok';
    if (delimiterCompliance === 'Wrong' || partsCountCompliance === 'Wrong' || partsCompliance === 'Wrong') {
        compliance = 'Wrong';
    }

    // Trim the trailing semicolon and space from details
    details = details.trim().replace(/; $/, '');

    return { compliance: compliance, details: details };
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

function downloadTemplate(templateName) {
    let templateData = [];

    switch (templateName) {
        case 'Template1':
            templateData = [
                ["Part", "Count", "Description", "Delimiter"],
                [1, 3, "Description", "_"],
                ["ExamplePart1", "", "", ""]
            ];
            break;
        case 'Template2':
            templateData = [
                ["Part", "Count", "Description", "Delimiter"],
                [2, 4, "Description", "-"],
                ["ExamplePart2", "", "", ""]
            ];
            break;
        case 'Template3':
            templateData = [
                ["Part", "Count", "Description", "Delimiter"],
                [3, 2, "Description", "."],
                ["ExamplePart3", "", "", ""]
            ];
            break;
        default:
            console.error('Unknown template name:', templateName);
            return;
    }

    const worksheet = XLSX.utils.aoa_to_sheet(templateData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'NamingConvention');
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });

    const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = templateName + ".xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
}
