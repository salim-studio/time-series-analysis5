/**
 * Functions to handle dummy variable creation and management for regression
 */

/**
 * Create configuration UI for dummy variables
 * @param {Array} dataset - The loaded dataset
 * @param {string} currentLang - Current language code
 * @returns {HTMLElement} - DOM element with the dummy variable interface
 */
export function createDummyVariableUI(dataset, currentLang) {
    if (!dataset || dataset.length === 0) return null;
    
    // Create container element
    const container = document.createElement('div');
    container.className = 'dummy-variables-section';
    
    // Get columns that might contain year/date information
    const columns = Object.keys(dataset[0]);
    const yearOptions = columns.map(col => 
        `<option value="${col}">${col}</option>`
    ).join('');
    
    // Labels for UI elements
    const labels = {
        dummyTitle: currentLang === 'ar' ? 'المتغيرات الوهمية' : 'Dummy Variables',
        yearColumn: currentLang === 'ar' ? 'عمود السنوات' : 'Year Column',
        selectYear: currentLang === 'ar' ? 'اختر السنة' : 'Select Year',
        selectQuarter: currentLang === 'ar' ? 'اختر الربع' : 'Select Quarter',
        addDummy: currentLang === 'ar' ? 'إضافة متغير وهمي' : 'Add Dummy',
        yearOption: currentLang === 'ar' ? 'سنة كاملة' : 'Full Year',
        quarterOption: currentLang === 'ar' ? 'ربع سنة' : 'Quarter',
        addYear: currentLang === 'ar' ? 'إضافة سنة' : 'Add Year',
        addQuarter: currentLang === 'ar' ? 'إضافة ربع' : 'Add Quarter',
        selectedDummies: currentLang === 'ar' ? 'المتغيرات الوهمية المحددة' : 'Selected Dummies'
    };
    
    // HTML for the dummy variable interface
    container.innerHTML = `
        <h4>${labels.dummyTitle}</h4>
        <div class="dummy-config">
            <div class="dummy-row">
                <label>${labels.yearColumn}:
                    <select id="yearColumnSelect">
                        <option value="">-- ${currentLang === 'ar' ? 'اختر' : 'Select'} --</option>
                        ${yearOptions}
                    </select>
                </label>
                <button id="loadYearsBtn">${currentLang === 'ar' ? 'تحميل السنوات' : 'Load Years'}</button>
            </div>
            
            <div id="yearSelectionContainer" class="year-selection hidden">
                <div class="dummy-row">
                    <select id="yearSelect" multiple>
                        <!-- Years will be added here -->
                    </select>
                    <select id="quarterSelect">
                        <option value="full">${labels.yearOption}</option>
                        <option value="1">Q1</option>
                        <option value="2">Q2</option>
                        <option value="3">Q3</option>
                        <option value="4">Q4</option>
                    </select>
                    <button id="addDummyBtn">${labels.addDummy}</button>
                </div>
                
                <div id="selectedDummiesContainer" class="selected-dummies">
                    <h5>${labels.selectedDummies}:</h5>
                    <ul id="selectedDummiesList">
                        <!-- Selected dummies will be listed here -->
                    </ul>
                </div>
            </div>
        </div>
    `;
    
    return container;
}

/**
 * Initialize the dummy variable functionality
 * @param {Array} dataset - The dataset to analyze
 * @param {string} currentLang - Current language code
 */
export function initDummyVariables(dataset, currentLang) {
    // Wait for the DOM to be ready
    setTimeout(() => {
        const loadYearsBtn = document.getElementById('loadYearsBtn');
        if (!loadYearsBtn) return;
        
        loadYearsBtn.addEventListener('click', () => loadYears(dataset));
        
        const addDummyBtn = document.getElementById('addDummyBtn');
        if (addDummyBtn) {
            addDummyBtn.addEventListener('click', () => addDummyVariable(dataset, currentLang));
        }
    }, 100);
}

/**
 * Load years from the selected column
 * @param {Array} dataset - The dataset to analyze
 */
function loadYears(dataset) {
    const yearColumnSelect = document.getElementById('yearColumnSelect');
    const yearColumn = yearColumnSelect.value;
    
    if (!yearColumn) return;
    
    // Extract unique years from the selected column
    const uniqueYears = new Set();
    
    dataset.forEach(row => {
        const value = row[yearColumn];
        if (value) {
            // Try to extract year from various formats
            let year;
            if (!isNaN(value)) {
                // If it's a number, use it directly
                year = Math.floor(parseFloat(value));
            } else if (typeof value === 'string') {
                // Try to extract year from string
                const yearMatch = value.match(/\b(19|20)\d{2}\b/);
                if (yearMatch) {
                    year = parseInt(yearMatch[0]);
                }
            }
            
            if (year && year >= 1900 && year <= 2100) {
                uniqueYears.add(year);
            }
        }
    });
    
    // Convert to array and sort
    const yearsArray = Array.from(uniqueYears).sort();
    
    // Populate the year select dropdown
    const yearSelect = document.getElementById('yearSelect');
    yearSelect.innerHTML = yearsArray.map(year => 
        `<option value="${year}">${year}</option>`
    ).join('');
    
    // Show the year selection container
    document.getElementById('yearSelectionContainer').classList.remove('hidden');
}

/**
 * Add a dummy variable based on selected year and quarter
 * @param {Array} dataset - The dataset to analyze
 * @param {string} currentLang - Current language code
 */
function addDummyVariable(dataset, currentLang) {
    const yearSelect = document.getElementById('yearSelect');
    const quarterSelect = document.getElementById('quarterSelect');
    const yearColumn = document.getElementById('yearColumnSelect').value;
    
    if (!yearColumn) return;
    
    // Get selected years
    const selectedYears = Array.from(yearSelect.selectedOptions).map(option => option.value);
    if (selectedYears.length === 0) return;
    
    const quarterValue = quarterSelect.value;
    const isFullYear = quarterValue === 'full';
    
    // Create dummy variables for each selected year
    selectedYears.forEach(year => {
        const dummyName = isFullYear 
            ? `dum_${year}` 
            : `dum_${year}_Q${quarterValue}`;
        
        // Add to the list of selected dummies
        addToSelectedDummies(dummyName, currentLang);
    });
}

/**
 * Add dummy variable to the selected list
 * @param {string} dummyName - Name of the dummy variable
 * @param {string} currentLang - Current language code
 */
function addToSelectedDummies(dummyName, currentLang) {
    const list = document.getElementById('selectedDummiesList');
    
    // Check if this dummy is already selected
    if (Array.from(list.children).some(li => li.dataset.name === dummyName)) {
        return;
    }
    
    const li = document.createElement('li');
    li.dataset.name = dummyName;
    li.innerHTML = `
        <span>${dummyName}</span>
        <button class="remove-dummy-btn">${currentLang === 'ar' ? 'حذف' : 'Remove'}</button>
    `;
    
    li.querySelector('.remove-dummy-btn').addEventListener('click', () => {
        list.removeChild(li);
    });
    
    list.appendChild(li);
}

/**
 * Get all selected dummy variables
 * @returns {Array} - Array of dummy variable names
 */
export function getSelectedDummies() {
    const list = document.getElementById('selectedDummiesList');
    if (!list) return [];
    
    return Array.from(list.children).map(li => li.dataset.name);
}

/**
 * Create dummy variables in the dataset for regression
 * @param {Array} data - Original data array
 * @param {string} yearColumn - Column containing year information
 * @param {Array} dummyNames - Array of dummy variable names to create
 * @returns {Object} - Extended dataset with dummy variables
 */
export function createDummyVariablesInData(data, yearColumn, dummyNames) {
    if (!data || !yearColumn || !dummyNames?.length) return data;
    
    // Clone the data
    const transformedData = JSON.parse(JSON.stringify(data));
    
    // Process each dummy variable
    dummyNames.forEach(dummyName => {
        // Parse the dummy variable name to extract year and quarter
        const match = dummyName.match(/dum_(\d{4})(?:_Q(\d))?/);
        if (!match) return;
        
        const year = match[1];
        const quarter = match[2] || null;
        
        // Add the dummy variable to each row
        transformedData.forEach(row => {
            const rowYearValue = row[yearColumn];
            let rowYear;
            let rowQuarter;
            
            // Extract year from the value
            if (typeof rowYearValue === 'number') {
                rowYear = Math.floor(rowYearValue);
                // Assume quarters are in decimal part (e.g. 2008.25 for Q1)
                const decimal = rowYearValue - rowYear;
                if (decimal > 0) {
                    rowQuarter = Math.ceil(decimal * 4);
                }
            } else if (typeof rowYearValue === 'string') {
                // Try to extract year from string
                const yearMatch = rowYearValue.match(/\b(19|20)\d{2}\b/);
                if (yearMatch) {
                    rowYear = parseInt(yearMatch[0]);
                }
                
                // Try to extract quarter from string
                const quarterMatch = rowYearValue.match(/Q(\d)/i);
                if (quarterMatch) {
                    rowQuarter = parseInt(quarterMatch[1]);
                }
            }
            
            // Set the dummy variable value
            if (rowYear === parseInt(year)) {
                if (quarter === null || rowQuarter === parseInt(quarter)) {
                    row[dummyName] = 1;
                } else {
                    row[dummyName] = 0;
                }
            } else {
                row[dummyName] = 0;
            }
        });
    });
    
    return transformedData;
}