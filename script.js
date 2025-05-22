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
const downloadGradesBtn = document.getElementById('downloadGradesBtn');
const dataPreviewSection = document.getElementById('dataPreviewSection');
const originalDataTable = document.getElementById('originalDataTable');
const firstMappedDataTable = document.getElementById('firstMappedDataTable');
const mappedDataTable = document.getElementById('mappedDataTable');
const finalGradesTable = document.getElementById('finalGradesTable');
const passingMarksSection = document.getElementById('passingMarksSection');
const calculateGradesBtn = document.getElementById('calculateGradesBtn');

/**
 * Global variables to store data
 * @type {Object}
 */
let originalJsonData = null;
let processedWorkbook = null;
let finalGradesData = null;

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
 * @param {Object} caps - Object containing upper cap values for each field
 * @param {boolean} includeGrades - Whether to include grades column
 */
function populateTable(data, tableBody, caps = null, includeGrades = false) {
    tableBody.innerHTML = '';

    // Calculate row totals if not already present
    const dataWithTotals = data.map(row => {
        if (row.Total === undefined) {
            const total = ['ESE', 'IA', 'CSE', 'TW', 'VIVA'].reduce((sum, field) => sum + (row[field] || 0), 0);
            return {...row, Total: Math.round(total * 100) / 100 };
        }
        return row;
    });

    // Determine headers based on whether grades are included
    const headers = includeGrades ? ['ESE', 'IA', 'CSE', 'TW', 'VIVA', 'Total', 'Grade'] : ['ESE', 'IA', 'CSE', 'TW', 'VIVA', 'Total'];

    // Find min and max values for numeric fields
    const stats = headers.reduce((acc, field) => {
        if (field !== 'Grade') {
            acc[field] = {
                min: Math.min(...dataWithTotals.map(row => row[field] || 0)),
                max: Math.max(...dataWithTotals.map(row => row[field] || 0))
            };
        }
        return acc;
    }, {});

    // Create table header row
    const headerRow = document.createElement('tr');
    headers.forEach(field => {
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
        headers.forEach(field => {
            const td = document.createElement('td');
            const value = row[field] || (field === 'Grade' ? '' : 0);
            td.textContent = value;

            // Add color coding for numeric fields
            if (field === 'Total') {
                if (value === stats.Total.max) {
                    td.style.backgroundColor = '#81c784';
                    td.title = 'Highest total score';
                    td.style.fontWeight = 'bold';
                }
            } else if (field === 'Grade') {
                // Color code grades
                const gradeColors = {
                    'O': '#4caf50', // Green
                    'A+': '#66bb6a', // Light Green
                    'A': '#81c784', // Lighter Green
                    'B+': '#fff176', // Yellow
                    'B': '#ffb74d', // Orange
                    'C': '#ff8a65', // Light Red
                    'D': '#e57373', // Red
                    'F': '#f44336' // Dark Red
                };
                if (gradeColors[value]) {
                    td.style.backgroundColor = gradeColors[value];
                    td.style.color = ['O', 'A+', 'A', 'D', 'F'].includes(value) ? 'white' : 'black';
                    td.style.fontWeight = 'bold';
                }
            } else if (field !== 'Grade') {
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

    // Add summary row at the bottom for numeric fields only
    if (!includeGrades) {
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

        // Reset processed workbook and final grades
        processedWorkbook = null;
        finalGradesData = null;

        // Update table with validation highlighting
        updateTableHighlighting();
        dataPreviewSection.style.display = 'flex';
        passingMarksSection.style.display = 'none';

        // Reset buttons
        downloadFirstBtn.disabled = true;
        downloadFinalBtn.disabled = true;
        downloadGradesBtn.disabled = true;
        calculateGradesBtn.disabled = true;

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

/**
 * Assign relative grades to passing students only
 * @param {Array} passingStudents - Array of students who passed ESE
 * @returns {Array} - Array with grades assigned
 */
function assignRelativeGrades(passingStudents) {
    if (passingStudents.length === 0) return [];

    // Calculate totals for passing students
    const totals = passingStudents.map(row =>
        (row.ESE || 0) + (row.IA || 0) + (row.CSE || 0) + (row.TW || 0) + (row.VIVA || 0)
    );

    // Pair totals with original index and sort descending
    const indexed = totals
        .map((t, i) => ({ total: t, idx: i }))
        .sort((a, b) => b.total - a.total);

    const N = passingStudents.length;

    // Compute "base" bucket size and remainder
    const base = Math.floor(N / 8);
    let rem = N - base * 8;

    // Define grade order and initial counts
    const grades = ['O', 'A+', 'A', 'B+', 'B', 'C', 'D', 'P']; // P for Pass (lowest passing grade)
    const counts = grades.reduce((acc, g) => {
        acc[g] = base;
        return acc;
    }, {});

    // Distribute any leftover students starting at the middle
    const middleOrder = ['B+', 'A', 'A+', 'O', 'D', 'C', 'B', 'P'];
    for (let i = 0; i < rem; i++) {
        counts[middleOrder[i]]++;
    }

    // Assign grade by slicing off each bucket from the top
    const cutoffs = {};
    let cursor = 0;
    for (const g of grades) {
        cutoffs[g] = [cursor, cursor + counts[g]]; // [start, end) in sorted list
        cursor += counts[g];
    }

    // Build result array
    const result = new Array(N);
    indexed.forEach((item, rank) => {
        // Find which grade bracket this rank falls into
        let assigned = 'P';
        for (const g of grades) {
            const [start, end] = cutoffs[g];
            if (rank >= start && rank < end) {
                assigned = g;
                break;
            }
        }
        result[item.idx] = {
            ...passingStudents[item.idx],
            Total: Math.round(item.total * 100) / 100,
            Grade: assigned
        };
    });

    return result;
}

/**
 * Calculate final grades with ESE passing marks logic
 */
function calculateFinalGrades() {
    if (!processedWorkbook || !processedWorkbook.final) {
        alert('Please convert the file first');
        return;
    }

    const esePassingPercentage = parseFloat(document.getElementById('esePassingMarks').value);
    if (isNaN(esePassingPercentage) || esePassingPercentage < 30 || esePassingPercentage > 70) {
        alert('Please enter valid ESE passing percentage (30-70%)');
        return;
    }

    // Convert percentage to actual marks out of 50
    const esePassingMarks = (esePassingPercentage / 100) * 50;

    const finalData = processedWorkbook.final;

    // Separate students based on ESE marks
    const passingStudents = [];
    const failingStudents = [];

    finalData.forEach((student, index) => {
        const studentWithTotal = {
            ...student,
            Total: Math.round(((student.ESE || 0) + (student.IA || 0) + (student.CSE || 0) + (student.TW || 0) + (student.VIVA || 0)) * 100) / 100,
            StudentIndex: index
        };

        if ((student.ESE || 0) >= esePassingMarks) {
            passingStudents.push(studentWithTotal);
        } else {
            failingStudents.push({
                ...studentWithTotal,
                Grade: 'F'
            });
        }
    });

    // Apply relative grading to passing students
    const gradedPassingStudents = assignRelativeGrades(passingStudents);

    // Combine all students back in original order
    const allStudentsWithGrades = new Array(finalData.length);

    // Place failing students
    failingStudents.forEach(student => {
        allStudentsWithGrades[student.StudentIndex] = student;
    });

    // Place passing students with grades
    gradedPassingStudents.forEach(student => {
        allStudentsWithGrades[student.StudentIndex] = student;
    });

    // Remove StudentIndex property
    finalGradesData = allStudentsWithGrades.map(({ StudentIndex, ...student }) => student);

    // Update statistics
    document.getElementById('totalStudents').textContent = finalData.length;
    document.getElementById('eseFailingCount').textContent = failingStudents.length;
    document.getElementById('gradingCount').textContent = passingStudents.length;

    // Update the final grades table
    populateTable(finalGradesData, finalGradesTable, FINAL_STAGE_MAPPINGS, true);

    // Enable download grades button
    downloadGradesBtn.disabled = false;
}

/**
 * Process Excel data and perform mappings
 */
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

        // Store data for downloads
        processedWorkbook = {
            first: firstStageData,
            final: finalStageData
        };

        // Enable download buttons
        downloadFirstBtn.disabled = false;
        downloadFinalBtn.disabled = false;
        calculateGradesBtn.disabled = false;

        // Show passing marks section
        passingMarksSection.style.display = 'block';

        // Update passing marks display immediately
        updatePassingMarksDisplay();

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
 * Download Excel file
 * @param {string} stage - Either 'first', 'final', or 'grades'
 */
function downloadExcel(stage) {
    let dataToDownload;
    let filename;

    if (stage === 'grades') {
        if (!finalGradesData) {
            alert('Please calculate final grades first');
            return;
        }
        dataToDownload = finalGradesData;
        filename = 'final_grades_with_ese_passing';
    } else {
        if (!processedWorkbook || !processedWorkbook[stage]) {
            alert('Please convert the file first');
            return;
        }
        dataToDownload = processedWorkbook[stage];
        filename = `${stage}_stage_mapping`;
    }

    // Generate timestamp for filename
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const wb = XLSX.utils.book_new();

    const ws = XLSX.utils.json_to_sheet(dataToDownload);
    XLSX.utils.book_append_sheet(wb, ws, `${stage.charAt(0).toUpperCase() + stage.slice(1)} Data`);
    XLSX.writeFile(wb, `${filename}_${timestamp}.xlsx`);
}

/**
 * Initialize the application
 */
function initializeApp() {
    // Reset all states
    originalJsonData = null;
    processedWorkbook = null;
    finalGradesData = null;

    // Reset download buttons
    if (downloadFirstBtn) downloadFirstBtn.disabled = true;
    if (downloadFinalBtn) downloadFinalBtn.disabled = true;
    if (downloadGradesBtn) downloadGradesBtn.disabled = true;
    if (calculateGradesBtn) calculateGradesBtn.disabled = true;

    // Hide sections
    if (dataPreviewSection) dataPreviewSection.style.display = 'none';
    if (passingMarksSection) passingMarksSection.style.display = 'none';
}

/**
 * Update table highlighting based on current data
 */
function updateTableHighlighting() {
    if (!originalJsonData) return;

    const caps = {
        ESE: parseFloat(document.getElementById('eseInput').value),
        IA: parseFloat(document.getElementById('iaInput').value),
        CSE: parseFloat(document.getElementById('cseInput').value),
        TW: parseFloat(document.getElementById('twInput').value),
        VIVA: parseFloat(document.getElementById('vivaInput').value)
    };

    // Update original data table
    populateTable(originalJsonData, originalDataTable, caps);

    // Update processed data tables if available
    if (processedWorkbook) {
        populateTable(processedWorkbook.first, firstMappedDataTable, FIRST_STAGE_MAPPINGS);
        populateTable(processedWorkbook.final, mappedDataTable, FINAL_STAGE_MAPPINGS);
    }

    // Update final grades table if available
    if (finalGradesData) {
        populateTable(finalGradesData, finalGradesTable, FINAL_STAGE_MAPPINGS, true);
    }
}

// Event Listeners
fileInput.addEventListener('change', handleFileUpload);
convertBtn.addEventListener('click', processExcel);
downloadFirstBtn.addEventListener('click', () => downloadExcel('first'));
downloadFinalBtn.addEventListener('click', () => downloadExcel('final'));
downloadGradesBtn.addEventListener('click', () => downloadExcel('grades'));
calculateGradesBtn.addEventListener('click', calculateFinalGrades);

// Add event listeners for cap input changes
['ese', 'ia', 'cse', 'tw', 'viva'].forEach(field => {
    document.getElementById(`${field}Input`).addEventListener('change', updateTableHighlighting);
});

// Add event listener for ESE passing marks change
document.getElementById('esePassingMarks').addEventListener('input', function() {
    updatePassingMarksDisplay();
    if (finalGradesData) {
        // Recalculate grades if data exists
        calculateFinalGrades();
    }
});

/**
 * Update the passing marks display and statistics
 */
function updatePassingMarksDisplay() {
    const esePassingPercentage = parseFloat(document.getElementById('esePassingMarks').value);

    if (!isNaN(esePassingPercentage) && processedWorkbook && processedWorkbook.final) {
        const esePassingMarks = (esePassingPercentage / 100) * 50;
        const finalData = processedWorkbook.final;

        // Count students
        const totalStudents = finalData.length;
        const failingStudents = finalData.filter(student => (student.ESE || 0) < esePassingMarks).length;
        const passingStudents = totalStudents - failingStudents;

        // Update display
        document.getElementById('totalStudents').textContent = totalStudents;
        document.getElementById('eseFailingCount').textContent = failingStudents;
        document.getElementById('gradingCount').textContent = passingStudents;

        // Show calculated passing marks
        const passingMarksDisplay = document.createElement('small');
        passingMarksDisplay.className = 'text-muted d-block mt-1';
        passingMarksDisplay.textContent = `Passing marks: ${esePassingMarks.toFixed(1)} out of 50`;

        // Remove existing display if any
        const existingDisplay = document.querySelector('.passing-marks-display');
        if (existingDisplay) {
            existingDisplay.remove();
        }

        // Add new display
        passingMarksDisplay.className += ' passing-marks-display';
        document.getElementById('esePassingMarks').parentNode.appendChild(passingMarksDisplay);
    } else {
        // Reset display if invalid input
        document.getElementById('totalStudents').textContent = '-';
        document.getElementById('eseFailingCount').textContent = '-';
        document.getElementById('gradingCount').textContent = '-';

        const existingDisplay = document.querySelector('.passing-marks-display');
        if (existingDisplay) {
            existingDisplay.remove();
        }
    }
}

// Initialize the app when the page loads
initializeApp();