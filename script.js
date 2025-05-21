/**
 * @param {string} filePath - Path to the file to be downloaded
 * @param {string} fileName - Optional custom name for the downloaded file
 */
function downloadFile(filePath, fileName = null) {
    try {

        const downloadLink = document.createElement('a');

        downloadLink.href = filePath;

        downloadLink.download = fileName || filePath.split('/').pop();

        document.body.appendChild(downloadLink);

        downloadLink.click();

        document.body.removeChild(downloadLink);

        console.log(`Download initiated for: ${filePath}`);
    } catch (error) {
        console.error("Error downloading file:", error);
        alert("There was an error downloading the file: " + error.message);
    }
}

// Add event listener to download link
document.addEventListener('DOMContentLoaded', function() {
    const downloadLinkElement = document.getElementById("downloadLink");

    if (downloadLinkElement) {
        downloadLinkElement.addEventListener("click", function() {

            const filePath = "./demo_tamplete.xlsx";

            const fileName = "Tamplete.xlsx";

            downloadFile(filePath, fileName);
        });
        console.log("Download link event listener added");
    } else {
        console.error("Download link element not found");
    }
});

// DOM Elements
const fileInput = document.getElementById('excelFile');
const convertBtn = document.getElementById('convertBtn');
const downloadFirstBtn = document.getElementById('downloadFirstBtn');
const downloadFinalBtn = document.getElementById('downloadFinalBtn');
const dataPreviewSection = document.getElementById('dataPreviewSection');
const originalDataTable = document.getElementById('originalDataTable');
const firstMappedDataTable = document.getElementById('firstMappedDataTable');
const mappedDataTable = document.getElementById('mappedDataTable');

/**
 * Global variables to store data
 * @type {Object}
 */
let originalJsonData = null;
let processedWorkbook = null;

// Constants for target mappings
const FIRST_STAGE_MAPPINGS = {
    ESE: 50,
    IA: 30,
    CSE: 20,
    TW: 25,
    VIVA: 25
};

const FINAL_STAGE_MAPPINGS = {
    ESE: 50,
    IA: 20,
    CSE: 10,
    TW: 10,
    VIVA: 10
};

/**
 * @param {Array} data - Array of objects containing the data
 * @param {HTMLElement} tableBody - Table body element to populate
 */
/**
 * @param {Array} data - Array of objects containing the data
 * @param {HTMLElement} tableBody - Table body element to populate
 * @param {Object} caps - Object containing upper cap values for each field
 */
function populateTable(data, tableBody, caps = null) {
    tableBody.innerHTML = '';

    // Calculate row totals
    const dataWithTotals = data.map(row => {
        const total = ['ESE', 'IA', 'CSE', 'TW', 'VIVA'].reduce((sum, field) => sum + (row[field] || 0), 0);
        return {...row, Total: total };
    });

    // Find min and max values for each field including Total
    const stats = ['ESE', 'IA', 'CSE', 'TW', 'VIVA', 'Total'].reduce((acc, field) => {
        acc[field] = {
            min: Math.min(...dataWithTotals.map(row => row[field] || 0)),
            max: Math.max(...dataWithTotals.map(row => row[field] || 0))
        };
        return acc;
    }, {});

    // Create table header row
    const headerRow = document.createElement('tr');
    ['ESE', 'IA', 'CSE', 'TW', 'VIVA', 'Total'].forEach(field => {
        const th = document.createElement('th');
        th.textContent = field;
        th.style.fontWeight = 'bold';
        th.style.backgroundColor = '#f5f5f5';
        headerRow.appendChild(th);
    });
    tableBody.appendChild(headerRow);

    // Create data rows
    dataWithTotals.forEach(row => {
        const tr = document.createElement('tr');
        ['ESE', 'IA', 'CSE', 'TW', 'VIVA', 'Total'].forEach(field => {
            const td = document.createElement('td');
            const value = row[field] || 0;
            td.textContent = value;

            // Add color coding
            if (field === 'Total') {
                if (value === stats.Total.max) {
                    td.style.backgroundColor = '#81c784';
                    td.title = 'Highest total score';
                    td.style.fontWeight = 'bold';
                }
            } else {
                if (value === 0) {
                    td.style.backgroundColor = '#ffebee'; // Light red for zero
                } else if (caps && value > caps[field]) {
                    td.style.backgroundColor = '#ff9e80'; // Orange for invalid (over cap)
                    td.title = `Value exceeds maximum cap of ${caps[field]}`;
                } else if (value < 0) {
                    td.style.backgroundColor = '#ff9e80'; // Orange for invalid (negative)
                    td.title = 'Negative values are not allowed';
                } else if (value === stats[field].max) {
                    td.style.backgroundColor = '#e8f5e9'; // Light green for max
                    td.title = 'Highest value in this column';
                } else if (caps && value === caps[field]) {
                    td.style.backgroundColor = '#e3f2fd'; // Light blue for perfect score
                    td.title = 'Maximum possible score';
                } else if (value === stats[field].min && value !== 0) {
                    td.style.backgroundColor = '#fff3e0'; // Light orange for min (excluding 0)
                    td.title = 'Lowest value in this column (excluding 0)';
                }
            }

            tr.appendChild(td);
        });
        tableBody.appendChild(tr);
    });

    // Add summary row at the bottom
    const summaryRow = document.createElement('tr');
    ['ESE', 'IA', 'CSE', 'TW', 'VIVA', 'Total'].forEach(field => {
        const td = document.createElement('td');
        const maxValue = stats[field].max;
        td.textContent = `Max: ${maxValue}`;
        td.style.backgroundColor = '#f5f5f5';
        td.style.fontWeight = 'bold';
        td.style.borderTop = '2px solid #ccc';
        summaryRow.appendChild(td);
    });
    tableBody.appendChild(summaryRow);
}

/**
 * @param {number} value - The input value to map
 * @param {number} fromMax - The maximum value of the input range
 * @param {number} toMax - The maximum value of the target range
 * @returns {number} - The mapped value
 */
function mapValue(value, fromMax, toMax) {
    return (value / fromMax) * toMax;
}

/**
 * @param {Event} e - Change event
 */
async function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        originalJsonData = XLSX.utils.sheet_to_json(worksheet);

        // Reset processed workbook
        processedWorkbook = null;

        // Update table with validation highlighting
        updateTableHighlighting();
        dataPreviewSection.style.display = 'flex';
        downloadFirstBtn.disabled = true;
        downloadFinalBtn.disabled = true;

        // Scroll to the data section
        dataPreviewSection.scrollIntoView({ behavior: 'smooth' });

        // Parse upper caps from column names if available
        const headers = Object.keys(originalJsonData[0] || {});
        headers.forEach(header => {
            const match = header.match(/([A-Z]+)\s*\((\d+)\)/);
            if (match) {
                const [, field, cap] = match;
                const inputId = field.toLowerCase() + 'Input';
                const input = document.getElementById(inputId);
                if (input) {
                    input.value = cap;
                }
            }
        });
    } catch (error) {
        console.error('Error reading file:', error);
        alert('Error reading file. Please make sure it\'s a valid Excel file.');
    }
}

/**
 * @param {Event} e - Click event
 */
/**
 * @param {string} header - Column header text
 * @returns {Object} Object containing field and cap value if found
 */
function parseHeaderCap(header) {
    const match = header.match(/([A-Z]+)\s*\((\d+)\)/);
    if (match) {
        const [, field, cap] = match;
        return { field, cap: parseInt(cap) };
    }
    return null;
}

async function processExcel(e) {
    e.preventDefault();

    if (!originalJsonData) {
        alert('Please upload an Excel file first');
        return;
    }

    // Get input values or parse from headers
    const headers = Object.keys(originalJsonData[0] || {});
    const caps = {
        ESE: parseFloat(document.getElementById('eseInput').value),
        IA: parseFloat(document.getElementById('iaInput').value),
        CSE: parseFloat(document.getElementById('cseInput').value),
        TW: parseFloat(document.getElementById('twInput').value),
        VIVA: parseFloat(document.getElementById('vivaInput').value)
    };

    // Try to get caps from headers if not manually entered
    headers.forEach(header => {
        const parsed = parseHeaderCap(header);
        if (parsed && (!caps[parsed.field] || isNaN(caps[parsed.field]))) {
            caps[parsed.field] = parsed.cap;
        }
    });

    // Validate all caps are present
    if (Object.values(caps).some(cap => !cap || isNaN(cap))) {
        alert('Please enter all upper cap values or ensure they are in column headers');
        return;
    }

    // Assign to variables for mapping
    const eseMax = caps.ESE;
    const iaMax = caps.IA;
    const cseMax = caps.CSE;
    const twMax = caps.TW;
    const vivaMax = caps.VIVA;

    try {
        // First stage mapping
        const firstStageData = originalJsonData.map(row => ({
            ESE: Math.round(mapValue(row.ESE || 0, eseMax, FIRST_STAGE_MAPPINGS.ESE) * 100) / 100,
            IA: Math.round(mapValue(row.IA || 0, iaMax, FIRST_STAGE_MAPPINGS.IA) * 100) / 100,
            CSE: Math.round(mapValue(row.CSE || 0, cseMax, FIRST_STAGE_MAPPINGS.CSE) * 100) / 100,
            TW: Math.round(mapValue(row.TW || 0, twMax, FIRST_STAGE_MAPPINGS.TW) * 100) / 100,
            VIVA: Math.round(mapValue(row.VIVA || 0, vivaMax, FIRST_STAGE_MAPPINGS.VIVA) * 100) / 100
        }));

        // Final stage mapping
        const finalStageData = firstStageData.map(row => ({
            ESE: Math.round(mapValue(row.ESE || 0, FIRST_STAGE_MAPPINGS.ESE, FINAL_STAGE_MAPPINGS.ESE) * 100) / 100,
            IA: Math.round(mapValue(row.IA || 0, FIRST_STAGE_MAPPINGS.IA, FINAL_STAGE_MAPPINGS.IA) * 100) / 100,
            CSE: Math.round(mapValue(row.CSE || 0, FIRST_STAGE_MAPPINGS.CSE, FINAL_STAGE_MAPPINGS.CSE) * 100) / 100,
            TW: Math.round(mapValue(row.TW || 0, FIRST_STAGE_MAPPINGS.TW, FINAL_STAGE_MAPPINGS.TW) * 100) / 100,
            VIVA: Math.round(mapValue(row.VIVA || 0, FIRST_STAGE_MAPPINGS.VIVA, FINAL_STAGE_MAPPINGS.VIVA) * 100) / 100
        }));

        // Create new worksheet with both mappings
        const newWorkbook = XLSX.utils.book_new();

        // Add first stage mapping sheet
        const firstStageSheet = XLSX.utils.json_to_sheet(firstStageData);
        XLSX.utils.book_append_sheet(newWorkbook, firstStageSheet, "First Stage Mapping");

        // Add final stage mapping sheet
        const finalStageSheet = XLSX.utils.json_to_sheet(finalStageData);
        XLSX.utils.book_append_sheet(newWorkbook, finalStageSheet, "Final Stage Mapping");

        // Store data for downloads
        processedWorkbook = {
            first: firstStageData,
            final: finalStageData
        };

        // Enable download buttons
        downloadFirstBtn.disabled = false;
        downloadFinalBtn.disabled = false;

        // Store processed data and update displays
        processedWorkbook = {
            first: firstStageData,
            final: finalStageData
        };

        // Update all tables with highlighting
        updateTableHighlighting();

        // Scroll to the data section
        dataPreviewSection.scrollIntoView({ behavior: 'smooth' });
    } catch (error) {
        console.error('Error processing file:', error);
        alert('Error processing file. Please make sure the file format is correct.');
    }
}

/**
 * @param {string} stage - Either 'first' or 'final'
 */
function downloadExcel(stage) {
    if (!processedWorkbook || !processedWorkbook[stage]) {
        alert('Please convert the file first');
        return;
    }

    // Generate timestamp for filename
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(processedWorkbook[stage]);
    XLSX.utils.book_append_sheet(wb, ws, `${stage.charAt(0).toUpperCase() + stage.slice(1)} Stage Mapping`);
    XLSX.writeFile(wb, `${stage}_stage_mapping_${timestamp}.xlsx`);
}

// Event Listeners
convertBtn.addEventListener('click', processExcel);

function initializeApp() {
    // Reset all states
    originalJsonData = null;
    processedWorkbook = null;

    // Reset download buttons
    if (downloadFirstBtn) downloadFirstBtn.disabled = true;
    if (downloadFinalBtn) downloadFinalBtn.disabled = true;

    // Hide preview section
    if (dataPreviewSection) dataPreviewSection.style.display = 'none';
}


function updateTableHighlighting() {
    if (!originalJsonData) return;

    const caps = {
        ESE: parseFloat(document.getElementById('eseInput').value),
        IA: parseFloat(document.getElementById('iaInput').value),
        CSE: parseFloat(document.getElementById('cseInput').value),
        TW: parseFloat(document.getElementById('twInput').value),
        VIVA: parseFloat(document.getElementById('vivaInput').value)
    };

    let hasInvalidData = false;
    originalJsonData.forEach(row => {
        Object.entries(row).forEach(([field, value]) => {
            if (value < 0 || (caps[field] && value > caps[field])) {
                hasInvalidData = true;
            }
        });
    });

    if (hasInvalidData) {
        // alert('Some values are invalid (negative or exceed maximum caps). These will be highlighted in orange.');
    }

    populateTable(originalJsonData, originalDataTable, caps);
    if (processedWorkbook) {
        populateTable(processedWorkbook.first, firstMappedDataTable, FIRST_STAGE_MAPPINGS);
        populateTable(processedWorkbook.final, mappedDataTable, FINAL_STAGE_MAPPINGS);
        updateHighestTotalMarks();
    }
}

// Event Listeners
fileInput.addEventListener('change', handleFileUpload);
convertBtn.addEventListener('click', processExcel);
downloadFirstBtn.addEventListener('click', () => downloadExcel('first'));
downloadFinalBtn.addEventListener('click', () => downloadExcel('final'));

// Add event listeners for cap input changes
['ese', 'ia', 'cse', 'tw', 'viva'].forEach(field => {
    document.getElementById(`${field}Input`).addEventListener('change', updateTableHighlighting);
});

// Initialize the app when the page loads
initializeApp();

function updatePassingMarks() {
    const input = document.getElementById("passingPercentageInput");
    const percentage = parseFloat(input.value);
    const highest = parseFloat(document.getElementById("highestTotalMarks").textContent);
    if (!isNaN(percentage) && percentage >= 35 && percentage <= 50 && !isNaN(highest)) {
        const passing = Math.round((percentage / 100) * highest * 100) / 100;
        document.getElementById("passingMarks").textContent = passing;
    } else {
        document.getElementById("passingMarks").textContent = "Invalid input";
    }
}


function updateHighestTotalMarks() {
    if (!processedWorkbook || !processedWorkbook.final) return;
    const totals = processedWorkbook.final.map(row => (
        (row.ESE || 0) + (row.IA || 0) + (row.CSE || 0) + (row.TW || 0) + (row.VIVA || 0)
    ));
    const maxTotal = Math.max(...totals);
    document.getElementById("highestTotalMarks").textContent = maxTotal.toFixed(2);
    updatePassingMarks();
}