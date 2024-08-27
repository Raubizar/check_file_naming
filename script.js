document.getElementById('folder-select-button').addEventListener('click', async () => {
    const directoryHandle = await window.showDirectoryPicker();
    const results = [];
    for await (const entry of directoryHandle.values()) {
        if (entry.kind === 'file') {
            const file = await entry.getFile();
            results.push({ name: file.name, compliance: 'Pending', details: 'Pending' });
        }
    }
    displayResults(results); // No longer passing processing time
});

document.getElementById('file-input').addEventListener('change', handleFileUpload);
document.getElementById('export-button').addEventListener('click', exportResults);
document.getElementById('excel-select-button').addEventListener('click', handleExcelSelection);

let namingConvention = null;
let fileNamesFromExcel = [];

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

async function handleExcelSelection() {
    try {
        const [fileHandle] = await window.showOpenFilePicker({
            types: [{
                description: 'Excel Files',
                accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
            }]
        });
        const file = await fileHandle.getFile();
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetSelect = document.getElementById('sheet-select');
        sheetSelect.innerHTML = ''; // Clear any previous options
        workbook.SheetNames.forEach((sheetName, index) => {
            const option = document.createElement('option');
            option.value = index;
            option.textContent = sheetName;
            sheetSelect.appendChild(option);
        });

        document.getElementById('excel-options').style.display = 'block';

        sheetSelect.addEventListener('change', function () {
            populateColumnSelect(workbook.Sheets[workbook.SheetNames[this.value]]);
        });

        // Load columns for the first sheet by default
        populateColumnSelect(workbook.Sheets[workbook.SheetNames[0]]);
    } catch (error) {
        console.error('Error selecting or reading Excel file:', error);
        alert('There was an issue selecting or reading the Excel file. Please try again.');
    }
}

function populateColumnSelect(sheet) {
    const columnSelect = document.getElementById('column-select');
    columnSelect.innerHTML = ''; // Clear any previous options

    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_col(C) + '1'; // Get the first row cell for header
        const cell = sheet[cellAddress] ? sheet[cellAddress].v : `Column ${C + 1}`;
        const option = document.createElement('option');
        option.value = C;
        option.textContent = cell;
        columnSelect.appendChild(option);
    }

    columnSelect.addEventListener('change', function () {
        loadFileNamesFromExcel(sheet, this.value);
    });

    // Automatically select the first column as default
    columnSelect.value = 0;
    loadFileNamesFromExcel(sheet, 0);
}

function loadFileNamesFromExcel(sheet, columnIndex) {
    try {
        fileNamesFromExcel = XLSX.utils.sheet_to_json(sheet, { header: 1 })
            .slice(1) // Skip the header row
            .map(row => row[columnIndex])
            .filter(name => typeof name === 'string' && name.trim() !== ''); // Filter out any empty or non-string cells

        console.log('File names from Excel loaded:', fileNamesFromExcel);

        displayResults(fileNamesFromExcel.map(name => ({ name, compliance: 'Pending', details: 'Pending' })));
    } catch (error) {
        console.error('Error loading file names from Excel:', error);
        alert('There was an issue loading file names from the selected Excel column. Please try again.');
    }
}

function displayResults(results) {
    const tbody = document.getElementById('results-table').querySelector('tbody');
    tbody.innerHTML = '';

    const thead = document.getElementById('results-table').querySelector('thead');
    thead.innerHTML = '';
    const headerRow = thead.insertRow();
    headerRow.insertCell(0).textContent = 'File Name';
    headerRow.insertCell(1).textContent = 'Compliance Status';
    headerRow.insertCell(2).textContent = 'Details';

    headerRow.cells[0].classList.add('header-cell');
    headerRow.cells[1].classList.add('header-cell');
    headerRow.cells[2].classList.add('header-cell');

    const totalFiles = results.length;
    let compliantCount = 0;

    const correctResults = results.filter(result => analyzeFileName(result.name).compliance === 'Ok');
    const incorrectResults = results.filter(result => analyzeFileName(result.name).compliance !== 'Ok');

    incorrectResults.forEach(result => {
        const row = tbody.insertRow();
        const analysis = analyzeFileName(result.name);
        row.insertCell(0).textContent = result.name;
        const complianceCell = row.insertCell(1);
        complianceCell.textContent = analysis.compliance;

        const detailsCell = row.insertCell(2);
        detailsCell.innerHTML = formatDetails(analysis.details, analysis.nonCompliantParts);

        if (analysis.compliance === 'Wrong') {
            complianceCell.style.color = 'red';
            row.cells[0].style.color = 'red';
        }
    });

    correctResults.forEach(result => {
        compliantCount++;
        const row = tbody.insertRow();
        const analysis = analyzeFileName(result.name);
        row.insertCell(0).textContent = result.name;
        row.insertCell(1).textContent = analysis.compliance;
        row.insertCell(2).innerHTML = formatDetails(analysis.details, analysis.nonCompliantParts);
    });

    // Calculate and update the summary
    const compliancePercentage = ((compliantCount / totalFiles) * 100).toFixed(2);
    document.getElementById('total-files').textContent = totalFiles;
    document.getElementById('names-comply').textContent = compliantCount;
    document.getElementById('compliance-percentage').textContent = `${compliancePercentage}%`;

    // Update the progress bar
    updateProgressBar(compliancePercentage);
}

function updateProgressBar(compliancePercentage) {
    const boxes = document.querySelectorAll('.progress-box');
    const filledBoxes = Math.floor(compliancePercentage / 10);
    const remainder = compliancePercentage % 10;

    boxes.forEach((box, index) => {
        if (index < filledBoxes) {
            box.classList.remove('yellow', 'red');
            box.classList.add('green');
        } else if (index === filledBoxes) {
            box.classList.remove('green', 'red');
            if (remainder > 0 && remainder < 10) {
                box.classList.add('yellow');
            } else {
                box.classList.add('red');
            }
        } else {
            box.classList.remove('green', 'yellow');
            box.classList.add('red');
        }
    });
}

function formatDetails(details, nonCompliantParts) {
    let formattedDetails = details;

    if (nonCompliantParts && nonCompliantParts.length > 0) {
        formattedDetails = formattedDetails.replace('Parts not compliant:', '<span class="error">Parts not compliant:</span>');
        nonCompliantParts.forEach(part => {
            const regex = new RegExp(`(${part})`, 'g');
            formattedDetails = formattedDetails.replace(regex, '<span class="error">$1</span>');
        });
    }

    return formattedDetails;
}

function analyzeFileName(fileName) {
    if (!namingConvention) {
        return { compliance: 'No naming convention uploaded', details: 'Please upload a naming convention file' };
    }

    let delimiterCompliance = 'Ok';
    let partsCountCompliance = 'Ok';
    let partsCompliance = 'Ok';
    let details = '';

    const dotPosition = fileName.lastIndexOf('.');
    if (dotPosition > 0) {
        fileName = fileName.substring(0, dotPosition);
    }

    const partsCount = parseInt(namingConvention[0][1], 10);
    const delimiter = namingConvention[0][3];
    const nameParts = fileName.split(delimiter);

    const expectedDelimiters = partsCount - 1;
    const actualDelimiters = (fileName.match(new RegExp(`\\${delimiter}`, 'g')) || []).length;
    if (actualDelimiters === expectedDelimiters) {
        details += 'Delimiter correct; ';
    } else {
        delimiterCompliance = 'Wrong';
        details += '<span class="error">Delimiter wrong</span>; ';
    }

    if (nameParts.length === partsCount) {
        details += 'Number of parts correct; ';
    } else {
        partsCountCompliance = 'Wrong';
        details += `<span class="error">Number of parts wrong (${nameParts.length})</span>; `;
    }

    let nonCompliantParts = [];
    for (let j = 0; j < nameParts.length; j++) {
        const allowedParts = namingConvention.slice(1).map(row => row[j]);
        let partAllowed = false;

        if (!allowedParts[0]) {
            nonCompliantParts.push(nameParts[j]);
            continue;
        }

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

    details = details.trim().replace(/; $/, '');

    return { compliance: compliance, details: details, nonCompliantParts: nonCompliantParts };
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
