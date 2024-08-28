document.getElementById('folder-select-button').addEventListener('click', async () => {
    if (!namingConvention) {
        alert("Please upload the Project Naming Convention first.");
        return;
    }
    const directoryHandle = await window.showDirectoryPicker();
    fileResultsFromFolder = []; // Clear previous results
    await traverseDirectory(directoryHandle, fileResultsFromFolder);
    displayResults(fileResultsFromFolder);
});

document.getElementById('file-input').addEventListener('change', handleFileUpload);
document.getElementById('excel-select-button').addEventListener('click', handleExcelSelection);

// Add the event listener for the export button here
document.getElementById('export-button').addEventListener('click', exportResults);

let namingConvention = null;
let fileNamesFromExcel = [];
let fileResultsFromFolder = [];  // New variable to store results from folder selection

async function traverseDirectory(directoryHandle, results, currentPath = '') {
    for await (const entry of directoryHandle.values()) {
        const fullPath = currentPath ? `${currentPath}/${entry.name}` : entry.name;
        if (entry.kind === 'file') {
            const file = await entry.getFile();
            results.push({
                name: file.name,
                path: currentPath, // Add the path to the results
                compliance: 'Pending',
                details: 'Pending'
            });
        } else if (entry.kind === 'directory') {
            await traverseDirectory(entry, results, fullPath); // Recursive call to traverse sub-directories
        }
    }
}

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

        // Re-analyze and display the results after loading the new naming convention
        if (fileNamesFromExcel.length > 0) {
            displayResults(fileNamesFromExcel.map(name => ({ name, compliance: 'Pending', details: 'Pending' })));
        } else if (fileResultsFromFolder.length > 0) {
            displayResults(fileResultsFromFolder);
        }
    } catch (error) {
        console.error('Error reading file:', error);
    }
}

document.getElementById('excel-select-button').addEventListener('click', handleExcelSelection);

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
    headerRow.insertCell(0).textContent = 'Folder Path';
    headerRow.insertCell(1).textContent = 'File Name';
    headerRow.insertCell(2).textContent = 'Compliance Status';
    headerRow.insertCell(3).textContent = 'Details';

    headerRow.cells[0].classList.add('header-cell');
    headerRow.cells[1].classList.add('header-cell');
    headerRow.cells[2].classList.add('header-cell');
    headerRow.cells[3].classList.add('header-cell');

    const folderGroups = groupByFolder(results);

    let totalFiles = 0;
    let compliantCount = 0;

    for (const [folder, files] of Object.entries(folderGroups)) {
        const folderRow = tbody.insertRow();
        folderRow.insertCell(0).textContent = folder;
        folderRow.insertCell(1).textContent = '';
        folderRow.insertCell(2).textContent = '';
        folderRow.insertCell(3).textContent = '';

        files.forEach(result => {
            const row = tbody.insertRow();
            row.insertCell(0).textContent = '';
            row.insertCell(1).textContent = result.name;

            const analysis = analyzeFileName(result.name);
            const complianceCell = row.insertCell(2);
            complianceCell.textContent = analysis.compliance;

            const detailsCell = row.insertCell(3);
            detailsCell.innerHTML = formatDetails(analysis.details, analysis.nonCompliantParts);

            if (analysis.compliance === 'Wrong') {
                complianceCell.style.color = 'red';
                row.cells[1].style.color = 'red';
            }

            if (analysis.compliance === 'Ok') {
                compliantCount++;
            }

            totalFiles++;
        });
    }

    // Calculate and update the summary
    const compliancePercentage = ((compliantCount / totalFiles) * 100).toFixed(2);
    document.getElementById('total-files').textContent = totalFiles;
    document.getElementById('names-comply').textContent = compliantCount;
    document.getElementById('compliance-percentage').textContent = `${compliancePercentage}%`;

    // Update the progress bar
    updateProgressBar(compliancePercentage);
}


function groupByFolder(results) {
    const folderGroups = {};
    results.forEach(result => {
        if (!folderGroups[result.path]) {
            folderGroups[result.path] = [];
        }
        folderGroups[result.path].push(result);
    });
    return folderGroups;
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

    // Add header row to the CSV data
    results.push(['Folder Path', 'File Name', 'Compliance Status', 'Details']);

    let currentFolderPath = '';

    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        const folderPath = cells[0].textContent.trim() || currentFolderPath; // Track the current folder path
        const fileName = cells[1].textContent.trim();
        const compliance = cells[2].textContent.trim();
        const details = cells[3].textContent.trim();

        // Update the current folder path only if a new one is found
        if (folderPath) {
            currentFolderPath = folderPath;
        }

        // Add the row to the results array
        results.push([currentFolderPath, fileName, compliance, details]);
    });

    // Convert the results array to CSV format
    const csvContent = "data:text/csv;charset=utf-8,"
        + results.map(e => e.map(cell => `"${cell}"`).join(",")).join("\n");

    // Create a downloadable link
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "results.csv");

    // Trigger the download
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

