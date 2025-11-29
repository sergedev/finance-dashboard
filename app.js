// Global variables to store data and charts
let transactions = [];
let categories = [];
let snapshots = [];
let charts = {};
let selectedCategories = new Set();
let currentTimeFilter = 'all';
let exchangeRates = { AUD: 0.50 }; // Fallback rates
let forecastMonths = 6; // Default forecast period
let uploadedFileData = null; // Store uploaded Excel file data

// Initialize the dashboard
document.addEventListener('DOMContentLoaded', () => {
    loadData();
    setupEventListeners();
    updateLastUpdatedTime();
});

// Load Excel data
async function loadData() {
    try {
        // Fetch exchange rates first
        await fetchExchangeRates();

        let arrayBuffer;
        let dataSource = 'dummy';

        // Check if user uploaded a file
        if (uploadedFileData) {
            console.log('✓ Loading uploaded data');
            arrayBuffer = uploadedFileData;
            dataSource = 'uploaded';
        } else {
            // Try to load from data/ folder first, fall back to data_dummy/ if not found
            let response;

            // Add cache-busting timestamp to ensure fresh data on every load
            const cacheBuster = `?t=${Date.now()}`;
            const fetchOptions = {
                cache: 'no-store',  // Don't use cached version
                headers: {
                    'Cache-Control': 'no-cache, no-store, must-revalidate',
                    'Pragma': 'no-cache',
                    'Expires': '0'
                }
            };

            try {
                response = await fetch(`data/finance.2025.xlsx${cacheBuster}`, fetchOptions);
                if (!response.ok) throw new Error('File not found');
                console.log('✓ Loading real data from data/finance.2025.xlsx');
                dataSource = 'real';
            } catch (e) {
                console.log('ℹ Loading dummy data from data_dummy/finance.2025.xlsx');
                response = await fetch(`data_dummy/finance.2025.xlsx${cacheBuster}`, fetchOptions);
                dataSource = 'dummy';
            }

            arrayBuffer = await response.arrayBuffer();
        }

        // Parse Excel file
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Extract data from each sheet
        const transactionsSheet = workbook.Sheets['Transactions'];
        const categoriesSheet = workbook.Sheets['Categories'];
        const snapshotsSheet = workbook.Sheets['Snapshots'];

        // Convert sheets to JSON
        const transactionsData = XLSX.utils.sheet_to_json(transactionsSheet);
        const categoriesData = XLSX.utils.sheet_to_json(categoriesSheet);
        const snapshotsData = XLSX.utils.sheet_to_json(snapshotsSheet);

        // Process transactions
        transactions = transactionsData.map(row => ({
            ...row,
            Amount: parseFloat(row.Amount) || 0,
            Balance: parseFloat(row.Balance) || 0,
            Date: parseExcelDate(row.Date)
        }));

        // Process categories
        categories = categoriesData;

        // Process snapshots
        snapshots = snapshotsData.map(row => ({
            ...row,
            Balance: parseFloat(row.Balance) || 0,
            Interest_Rate: parseFloat(row.Interest_Rate) || 0,
            Date: parseExcelDate(row.Date)
        }));

        // Initialize both dashboards
        initializeDashboard();
        initializeNetWorth();

        // Update data source banner
        updateDataSourceBanner(dataSource);

    } catch (error) {
        console.error('Error loading data:', error);
        alert('Error loading data. Please ensure you have uploaded a file or that finance.2025.xlsx is in the data/ or data_dummy/ folder.');
    }
}

// Update the data source banner
function updateDataSourceBanner(source) {
    const banner = document.getElementById('dataSourceText');
    const resetBtn = document.getElementById('resetToDemo');

    if (source === 'uploaded') {
        banner.textContent = '✓ Using your uploaded data';
        resetBtn.style.display = 'block';
    } else if (source === 'real') {
        banner.textContent = '✓ Using real data from data/ folder';
        resetBtn.style.display = 'none';
    } else {
        banner.textContent = '📊 Using demo data - Upload your file to get started';
        resetBtn.style.display = 'none';
    }
}

// Handle file upload
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    // Verify it's an Excel file
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        alert('Please upload an Excel file (.xlsx or .xls)');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            uploadedFileData = e.target.result;
            console.log('✓ File uploaded successfully:', file.name);
            loadData();
        } catch (error) {
            console.error('Error reading file:', error);
            alert('Error reading the Excel file. Please make sure it has the correct format.');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Make table sortable by clicking column headers
function makeSortable(tableId) {
    const table = document.getElementById(tableId);
    if (!table) return;

    const headers = table.querySelectorAll('th.sortable');

    headers.forEach((header, columnIndex) => {
        let currentSort = null; // null, 'asc', or 'desc'

        header.addEventListener('click', () => {
            // Toggle sort direction
            if (currentSort === 'asc') {
                currentSort = 'desc';
            } else if (currentSort === 'desc') {
                currentSort = 'asc';
            } else {
                currentSort = 'asc';
            }

            // Remove sort classes from all headers
            headers.forEach(h => h.classList.remove('sort-asc', 'sort-desc'));

            // Add sort class to clicked header
            header.classList.add(currentSort === 'asc' ? 'sort-asc' : 'sort-desc');

            // Sort the table
            sortTable(table, columnIndex, currentSort);
        });
    });
}

// Sort table by column index
function sortTable(table, columnIndex, direction) {
    const tbody = table.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));

    rows.sort((a, b) => {
        const aCell = a.children[columnIndex];
        const bCell = b.children[columnIndex];

        // Use data-value if available, otherwise use text content
        let aValue = aCell.dataset.value !== undefined ? aCell.dataset.value : aCell.textContent.trim();
        let bValue = bCell.dataset.value !== undefined ? bCell.dataset.value : bCell.textContent.trim();

        // Try to parse as numbers
        const aNum = parseFloat(aValue.replace(/[£,]/g, ''));
        const bNum = parseFloat(bValue.replace(/[£,]/g, ''));

        // Compare as numbers if both are valid numbers
        if (!isNaN(aNum) && !isNaN(bNum)) {
            return direction === 'asc' ? aNum - bNum : bNum - aNum;
        }

        // Otherwise compare as strings
        if (direction === 'asc') {
            return aValue.localeCompare(bValue);
        } else {
            return bValue.localeCompare(aValue);
        }
    });

    // Re-append rows in sorted order
    rows.forEach(row => tbody.appendChild(row));
}

// Parse amount/balance values - handles £ symbols, commas, and whitespace
function parseAmount(amountStr) {
    if (!amountStr) return 0;
    // Remove £ symbol, commas, quotes, and whitespace, then parse
    const cleaned = amountStr.toString().replace(/[£,"\s]/g, '');
    return parseFloat(cleaned) || 0;
}

// Parse Excel date - handles Excel serial numbers and text dates
function parseExcelDate(dateValue) {
    if (!dateValue) return new Date();

    // If it's already a Date object, return it
    if (dateValue instanceof Date) return dateValue;

    // If it's a number (Excel serial date)
    if (typeof dateValue === 'number') {
        // Excel dates are days since 1900-01-01 (with leap year bug adjustment)
        const excelEpoch = new Date(1899, 11, 30);
        return new Date(excelEpoch.getTime() + dateValue * 86400000);
    }

    // If it's a string, use the parseDate function
    return parseDate(dateValue.toString());
}

// Parse date - handles both DD-MMM-YY and DD/MM/YYYY formats
function parseDate(dateStr) {
    if (!dateStr) return new Date();

    // Check if it's DD/MM/YYYY format (contains /)
    if (dateStr.includes('/')) {
        const parts = dateStr.split('/');
        const day = parseInt(parts[0]);
        const month = parseInt(parts[1]) - 1; // Month is 0-indexed
        const year = parseInt(parts[2]);
        return new Date(year, month, day);
    }

    // Otherwise parse DD-MMM-YY format (contains -)
    const months = {
        'Jan': 0, 'Feb': 1, 'Mar': 2, 'Apr': 3, 'May': 4, 'Jun': 5,
        'Jul': 6, 'Aug': 7, 'Sep': 8, 'Oct': 9, 'Nov': 10, 'Dec': 11
    };

    const parts = dateStr.split('-');
    const day = parseInt(parts[0]);
    const month = months[parts[1]];
    const year = parts[2].length === 2 ? 2000 + parseInt(parts[2]) : parseInt(parts[2]);

    return new Date(year, month, day);
}

// Initialize dashboard after data is loaded
function initializeDashboard() {
    // Get unique categories from BOTH transactions AND categories.csv
    const categoriesFromCSV = [...new Set(categories.map(c => c.Category))].filter(cat => cat);
    const categoriesFromTransactions = [...new Set(transactions.map(t => t.Category))].filter(cat => cat);

    // Merge and deduplicate both sources
    const uniqueCategories = [...new Set([...categoriesFromCSV, ...categoriesFromTransactions])].sort();

    // Initialize all categories as selected except Transfers
    uniqueCategories.forEach(cat => {
        if (cat !== 'Transfers') {
            selectedCategories.add(cat);
        }
    });

    // Create category checkboxes
    createCategoryCheckboxes(uniqueCategories);

    // Update dashboard
    updateDashboard();
}

// Create category filter checkboxes
function createCategoryCheckboxes(categories) {
    const container = document.getElementById('categoryCheckboxes');
    container.innerHTML = '';

    categories.sort().forEach(category => {
        const div = document.createElement('div');
        div.className = 'checkbox-item';

        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = `cat-${category}`;
        checkbox.value = category;
        checkbox.checked = selectedCategories.has(category);
        checkbox.addEventListener('change', handleCategoryChange);

        const label = document.createElement('label');
        label.htmlFor = `cat-${category}`;
        label.textContent = category;

        div.appendChild(checkbox);
        div.appendChild(label);
        container.appendChild(div);
    });
}

// Setup event listeners
function setupEventListeners() {
    // File upload
    document.getElementById('fileUpload').addEventListener('change', handleFileUpload);

    // Reset to demo data
    document.getElementById('resetToDemo').addEventListener('click', () => {
        uploadedFileData = null;
        loadData();
    });

    // Time filter - Month select
    document.getElementById('monthSelect').addEventListener('change', (e) => {
        currentTimeFilter = e.target.value;
        // Clear preset button active states
        document.querySelectorAll('.preset-btn').forEach(btn => btn.classList.remove('active'));
        updateDashboard();
    });

    // Time filter - Preset buttons
    document.querySelectorAll('.preset-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const preset = e.target.dataset.preset;
            currentTimeFilter = preset;

            // Update active state
            document.querySelectorAll('.preset-btn').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');

            // Reset month select
            document.getElementById('monthSelect').value = 'all';

            updateDashboard();
        });
    });

    // Toggle all categories
    document.getElementById('toggleAll').addEventListener('click', () => {
        const allChecked = selectedCategories.size > 0;

        if (allChecked) {
            // Deselect all
            selectedCategories.clear();
            document.getElementById('toggleAll').textContent = 'Select All';
        } else {
            // Select all - get all visible checkboxes
            document.querySelectorAll('#categoryCheckboxes input[type="checkbox"]').forEach(cb => {
                selectedCategories.add(cb.value);
            });
            document.getElementById('toggleAll').textContent = 'Deselect All';
        }

        // Update all checkboxes
        document.querySelectorAll('#categoryCheckboxes input[type="checkbox"]').forEach(cb => {
            cb.checked = selectedCategories.has(cb.value);
        });

        updateDashboard();
    });
}

// Handle category checkbox change
function handleCategoryChange(e) {
    const category = e.target.value;

    if (e.target.checked) {
        selectedCategories.add(category);
    } else {
        selectedCategories.delete(category);
    }

    // Update toggle button text
    document.getElementById('toggleAll').textContent =
        selectedCategories.size > 0 ? 'Deselect All' : 'Select All';

    updateDashboard();
}

// Filter transactions based on current filters
function getFilteredTransactions() {
    return transactions.filter(t => {
        // Category filter
        if (!selectedCategories.has(t.Category)) {
            return false;
        }

        // Time filter
        const date = t.Date;
        const now = new Date();

        switch (currentTimeFilter) {
            case '30days':
                const thirtyDaysAgo = new Date(now);
                thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
                return date >= thirtyDaysAgo;

            case '6months':
                const sixMonthsAgo = new Date(now);
                sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
                return date >= sixMonthsAgo;

            case 'ytd':
                const startOfYear = new Date(now.getFullYear(), 0, 1);
                return date >= startOfYear;

            case 'all':
                return true;

            default:
                // Specific month (format: 2025-01)
                if (currentTimeFilter.includes('-')) {
                    const [year, month] = currentTimeFilter.split('-').map(Number);
                    return date.getFullYear() === year && date.getMonth() === month - 1;
                }
                return true;
        }
    });
}

// Calculate summary statistics
function calculateStats(filteredTransactions) {
    // Exclude Transfers from expense calculations
    const transactionsForCalc = filteredTransactions.filter(t => t.Category !== 'Transfers');

    const income = transactionsForCalc
        .filter(t => t.Amount > 0)
        .reduce((sum, t) => sum + t.Amount, 0);

    const expenses = Math.abs(transactionsForCalc
        .filter(t => t.Amount < 0)
        .reduce((sum, t) => sum + t.Amount, 0));

    const netSavings = income - expenses;
    const savingsRate = income > 0 ? (netSavings / income * 100) : 0;

    // Calculate average monthly spending
    const monthsSet = new Set();
    transactionsForCalc.forEach(t => {
        const monthKey = `${t.Date.getFullYear()}-${t.Date.getMonth()}`;
        monthsSet.add(monthKey);
    });
    const monthCount = monthsSet.size || 1;
    const avgMonthly = expenses / monthCount;

    return {
        income,
        expenses,
        netSavings,
        savingsRate,
        avgMonthly
    };
}

// Format number with thousand separators
function formatCurrency(amount) {
    return amount.toLocaleString('en-GB', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

// Update summary stats display
function updateSummaryStats(stats) {
    document.getElementById('totalIncome').textContent = `£${formatCurrency(stats.income)}`;
    document.getElementById('totalExpenses').textContent = `£${formatCurrency(stats.expenses)}`;
    document.getElementById('netSavings').textContent = `£${formatCurrency(stats.netSavings)}`;
    document.getElementById('savingsRate').textContent = `${stats.savingsRate.toFixed(1)}%`;
    document.getElementById('avgMonthly').textContent = `£${formatCurrency(stats.avgMonthly)}`;
}

// Prepare data for monthly trend chart
function getMonthlyTrendData(filteredTransactions) {
    // Exclude Transfers
    const transactionsForChart = filteredTransactions.filter(t => t.Category !== 'Transfers');

    const monthlyData = {};

    transactionsForChart.forEach(t => {
        const monthKey = `${t.Date.getFullYear()}-${String(t.Date.getMonth() + 1).padStart(2, '0')}`;

        if (!monthlyData[monthKey]) {
            monthlyData[monthKey] = { income: 0, expenses: 0 };
        }

        if (t.Amount > 0) {
            monthlyData[monthKey].income += t.Amount;
        } else {
            monthlyData[monthKey].expenses += Math.abs(t.Amount);
        }
    });

    // Sort by date
    const sortedMonths = Object.keys(monthlyData).sort();

    const labels = sortedMonths.map(month => {
        const [year, monthNum] = month.split('-');
        const date = new Date(year, monthNum - 1);
        return date.toLocaleDateString('en-GB', { month: 'short', year: '2-digit' });
    });

    const incomeData = sortedMonths.map(month => monthlyData[month].income);
    const expenseData = sortedMonths.map(month => monthlyData[month].expenses);

    return { labels, incomeData, expenseData };
}

// Prepare data for category charts
function getCategoryData(filteredTransactions) {
    // Exclude Income and Transfers for expense breakdown
    const expenseTransactions = filteredTransactions.filter(t =>
        t.Amount < 0 && t.Category !== 'Income' && t.Category !== 'Transfers'
    );

    const categoryTotals = {};

    expenseTransactions.forEach(t => {
        if (!categoryTotals[t.Category]) {
            categoryTotals[t.Category] = 0;
        }
        categoryTotals[t.Category] += Math.abs(t.Amount);
    });

    // Sort by amount
    const sortedCategories = Object.entries(categoryTotals)
        .sort((a, b) => b[1] - a[1]);

    const labels = sortedCategories.map(([cat]) => cat);
    const data = sortedCategories.map(([, amount]) => amount);

    return { labels, data };
}

// Generate colors for charts
function generateColors(count) {
    const colors = [
        '#ef4444', '#f59e0b', '#10b981', '#3b82f6', '#8b5cf6',
        '#ec4899', '#14b8a6', '#f97316', '#06b6d4', '#84cc16',
        '#6366f1', '#f43f5e', '#22d3ee', '#a855f7', '#eab308'
    ];

    const result = [];
    for (let i = 0; i < count; i++) {
        result.push(colors[i % colors.length]);
    }
    return result;
}

// Create or update monthly trend chart
function updateMonthlyTrendChart(data) {
    const ctx = document.getElementById('monthlyTrendChart').getContext('2d');

    if (charts.monthlyTrend) {
        charts.monthlyTrend.destroy();
    }

    charts.monthlyTrend = new Chart(ctx, {
        type: 'line',
        data: {
            labels: data.labels,
            datasets: [
                {
                    label: 'Income',
                    data: data.incomeData,
                    borderColor: '#10b981',
                    backgroundColor: 'rgba(16, 185, 129, 0.1)',
                    borderWidth: 3,
                    tension: 0.4,
                    fill: true
                },
                {
                    label: 'Expenses',
                    data: data.expenseData,
                    borderColor: '#ef4444',
                    backgroundColor: 'rgba(239, 68, 68, 0.1)',
                    borderWidth: 3,
                    tension: 0.4,
                    fill: true
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            animation: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'top'
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    callbacks: {
                        label: function(context) {
                            return `${context.dataset.label}: £${context.parsed.y.toFixed(2)}`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return '£' + value.toFixed(0);
                        }
                    }
                }
            }
        }
    });
}

// Create or update category pie chart
function updateCategoryPieChart(data) {
    const ctx = document.getElementById('categoryPieChart').getContext('2d');

    if (charts.categoryPie) {
        charts.categoryPie.destroy();
    }

    if (data.labels.length === 0) {
        return;
    }

    charts.categoryPie = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: data.labels,
            datasets: [{
                data: data.data,
                backgroundColor: generateColors(data.labels.length),
                borderWidth: 2,
                borderColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            animation: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'right'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const label = context.label || '';
                            const value = context.parsed;
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const percentage = ((value / total) * 100).toFixed(1);
                            return `${label}: £${value.toFixed(2)} (${percentage}%)`;
                        }
                    }
                }
            }
        }
    });
}

// Create or update category bar chart
function updateCategoryBarChart(data) {
    const ctx = document.getElementById('categoryBarChart').getContext('2d');

    if (charts.categoryBar) {
        charts.categoryBar.destroy();
    }

    if (data.labels.length === 0) {
        return;
    }

    // Take top 10 categories for better visibility
    const topLabels = data.labels.slice(0, 10);
    const topData = data.data.slice(0, 10);

    charts.categoryBar = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: topLabels,
            datasets: [{
                label: 'Spending',
                data: topData,
                backgroundColor: generateColors(topLabels.length),
                borderWidth: 0
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: true,
            animation: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return `£${context.parsed.x.toFixed(2)}`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return '£' + value.toFixed(0);
                        }
                    }
                }
            }
        }
    });
}

// Main update function
function updateDashboard() {
    const filtered = getFilteredTransactions();

    // Update summary stats
    const stats = calculateStats(filtered);
    updateSummaryStats(stats);

    // Update charts
    const monthlyData = getMonthlyTrendData(filtered);
    updateMonthlyTrendChart(monthlyData);

    const categoryData = getCategoryData(filtered);
    updateCategoryPieChart(categoryData);
    updateCategoryBarChart(categoryData);
}

// Update last updated time
function updateLastUpdatedTime() {
    const now = new Date();
    const formatted = now.toLocaleDateString('en-GB', {
        day: '2-digit',
        month: 'short',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
    document.getElementById('lastUpdated').textContent = formatted;
}

// ========================================
// NET WORTH & FORECAST FUNCTIONS
// ========================================

// Fetch exchange rates from API
async function fetchExchangeRates() {
    try {
        const response = await fetch('https://api.exchangerate-api.com/v4/latest/GBP');
        const data = await response.json();
        if (data && data.rates) {
            exchangeRates.AUD = 1 / data.rates.AUD; // Convert to GBP from AUD
        }
    } catch (error) {
        console.warn('Failed to fetch exchange rates, using fallback:', error);
        // Fallback rates already set in global variables
    }
}

// Convert foreign currency to GBP
function convertToGBP(amount, currency) {
    if (currency === 'GBP') return amount;
    return amount * (exchangeRates[currency] || 0.50);
}

// Initialize net worth section
function initializeNetWorth() {
    if (snapshots.length === 0) {
        console.warn('No snapshot data available');
        return;
    }

    // Update all sections
    updateNetWorthSummary();
    updateHoldingsTable();
    updateHistoricalChart();
    updateForecast();

    // Setup event listeners
    setupNetWorthEventListeners();

    // Set default forecast date (6 months from now)
    const defaultDate = new Date();
    defaultDate.setMonth(defaultDate.getMonth() + 6);
    document.getElementById('forecastDate').valueAsDate = defaultDate;
}

// Get current snapshot (most recent date for each account)
function getCurrentSnapshot() {
    const accountSnapshots = {};

    snapshots.forEach(snap => {
        const key = snap.Account_Name;
        if (!accountSnapshots[key] || snap.Date > accountSnapshots[key].Date) {
            accountSnapshots[key] = snap;
        }
    });

    // Filter out closed accounts (Balance = 0)
    return Object.values(accountSnapshots).filter(acc => acc.Balance > 0);
}

// Update net worth summary cards
function updateNetWorthSummary() {
    const current = getCurrentSnapshot();

    let totalNetWorth = 0;
    let totalISAs = 0;
    let totalLiquid = 0;
    let foreignTotal = 0;
    let foreignGBP = 0;

    current.forEach(acc => {
        const gbpValue = convertToGBP(acc.Balance, acc.Currency);
        totalNetWorth += gbpValue;

        // ISA calculation
        if (acc.Account_Name.toUpperCase().includes('ISA')) {
            totalISAs += gbpValue;
        } else {
            totalLiquid += gbpValue;
        }

        // Foreign holdings
        if (acc.Currency !== 'GBP') {
            foreignTotal += acc.Balance;
            foreignGBP += gbpValue;
        }
    });

    document.getElementById('totalNetWorth').textContent = `£${formatCurrency(totalNetWorth)}`;
    document.getElementById('totalISAs').textContent = `£${formatCurrency(totalISAs)}`;
    document.getElementById('totalLiquid').textContent = `£${formatCurrency(totalLiquid)}`;
    document.getElementById('foreignHoldings').textContent =
        `${formatCurrency(foreignTotal)} AUD (£${formatCurrency(foreignGBP)})`;
}

// Update holdings table
function updateHoldingsTable() {
    const current = getCurrentSnapshot();

    // Sort by GBP equivalent (descending)
    current.sort((a, b) => {
        const aGBP = convertToGBP(a.Balance, a.Currency);
        const bGBP = convertToGBP(b.Balance, b.Currency);
        return bGBP - aGBP;
    });

    const tbody = document.querySelector('#holdingsTable tbody');
    tbody.innerHTML = '';

    current.forEach(acc => {
        const gbpValue = convertToGBP(acc.Balance, acc.Currency);
        const row = document.createElement('tr');

        const interestDisplay = acc.Interest_Rate
            ? `${acc.Interest_Rate.toFixed(2)}% ${acc.Rate_Type || ''}`
            : '';

        const lastUpdated = acc.Date
            ? acc.Date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit' })
            : '';

        row.innerHTML = `
            <td>${acc.Account_Name}</td>
            <td data-value="${acc.Balance}">${formatCurrency(acc.Balance)}</td>
            <td>${acc.Currency}</td>
            <td data-value="${gbpValue}">£${formatCurrency(gbpValue)}</td>
            <td data-value="${acc.Interest_Rate || 0}">${interestDisplay}</td>
            <td data-value="${acc.Date ? acc.Date.getTime() : 0}">${lastUpdated}</td>
        `;

        tbody.appendChild(row);
    });
}

// Update historical net worth chart
function updateHistoricalChart() {
    // Group snapshots by date and calculate total net worth
    const dateGroups = {};

    snapshots.forEach(snap => {
        if (snap.Balance === 0) return; // Skip closed accounts

        const dateKey = snap.Date.toISOString().split('T')[0];
        if (!dateGroups[dateKey]) {
            dateGroups[dateKey] = { date: snap.Date, total: 0 };
        }

        const gbpValue = convertToGBP(snap.Balance, snap.Currency);
        dateGroups[dateKey].total += gbpValue;
    });

    // Sort by date
    const sortedData = Object.values(dateGroups).sort((a, b) => a.date - b.date);

    const labels = sortedData.map(d => d.date.toLocaleDateString('en-GB', {
        month: 'short',
        year: '2-digit'
    }));
    const data = sortedData.map(d => d.total);

    // Calculate growth stats
    if (sortedData.length >= 2) {
        const firstValue = sortedData[0].total;
        const lastValue = sortedData[sortedData.length - 1].total;
        const totalGrowth = lastValue - firstValue;

        const firstDate = sortedData[0].date;
        const lastDate = sortedData[sortedData.length - 1].date;
        const monthsDiff = (lastDate - firstDate) / (1000 * 60 * 60 * 24 * 30.44);
        const avgMonthlyGrowth = monthsDiff > 0 ? totalGrowth / monthsDiff : 0;

        document.getElementById('totalGrowth').textContent =
            `£${formatCurrency(Math.abs(totalGrowth))} ${totalGrowth >= 0 ? 'increase' : 'decrease'}`;
        document.getElementById('avgMonthlyGrowth').textContent =
            `£${formatCurrency(Math.abs(avgMonthlyGrowth))}/mo`;
    }

    // Create chart
    const ctx = document.getElementById('networthHistoryChart').getContext('2d');
    if (charts.networthHistory) {
        charts.networthHistory.destroy();
    }

    charts.networthHistory = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Net Worth (GBP)',
                data: data,
                borderColor: '#10b981',
                backgroundColor: 'rgba(16, 185, 129, 0.1)',
                borderWidth: 3,
                tension: 0.4,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            animation: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return `£${formatCurrency(context.parsed.y)}`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: false,
                    ticks: {
                        callback: function(value) {
                            return '£' + (value / 1000).toFixed(0) + 'k';
                        }
                    }
                }
            }
        }
    });
}

// Calculate forecast with compound interest
function calculateForecast() {
    const current = getCurrentSnapshot();
    const months = forecastMonths;
    const monthlyContribution = parseFloat(document.getElementById('monthlyContribution').value) || 0;

    const accountForecasts = current.map(acc => {
        const principal = convertToGBP(acc.Balance, acc.Currency);
        const annualRate = acc.Interest_Rate / 100;

        // Calculate monthly rate based on rate type
        let monthlyRate;
        if (acc.Rate_Type === 'AER') {
            // AER already accounts for compounding - convert to monthly equivalent
            monthlyRate = Math.pow(1 + annualRate, 1/12) - 1;
        } else {
            // Gross and PA - use simple division (nominal rate compounded monthly)
            monthlyRate = annualRate / 12;
        }

        // Compound interest formula: FV = PV * (1 + r)^n
        let futureValue = principal * Math.pow(1 + monthlyRate, months);

        // Add monthly contributions (simplified: added at end of each month)
        if (monthlyContribution > 0) {
            // Distribute contribution across accounts proportionally
            const proportion = principal / current.reduce((sum, a) =>
                sum + convertToGBP(a.Balance, a.Currency), 0);
            const accountContribution = monthlyContribution * proportion;

            // Future value of annuity: FV = PMT * [(1 + r)^n - 1] / r
            if (monthlyRate > 0) {
                futureValue += accountContribution * (Math.pow(1 + monthlyRate, months) - 1) / monthlyRate;
            } else {
                futureValue += accountContribution * months;
            }
        }

        return {
            name: acc.Account_Name,
            current: principal,
            projected: futureValue,
            growth: futureValue - principal
        };
    });

    return accountForecasts;
}

// Update forecast display
function updateForecast() {
    const forecasts = calculateForecast();
    const totalProjected = forecasts.reduce((sum, f) => sum + f.projected, 0);

    // Update forecast card
    document.getElementById('forecastAmount').textContent = `£${formatCurrency(totalProjected)}`;

    const timeframeText = forecastMonths === 6 ? '6 months' :
                         forecastMonths === 12 ? '1 year' :
                         forecastMonths === 24 ? '2 years' :
                         `${forecastMonths} months`;
    document.getElementById('forecastTimeframe').textContent = `in ${timeframeText}`;

    // Update forecast breakdown table
    const tbody = document.querySelector('#forecastTable tbody');
    tbody.innerHTML = '';

    forecasts.forEach(f => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td data-value="${f.name}">${f.name}</td>
            <td data-value="${f.current}">£${formatCurrency(f.current)}</td>
            <td data-value="${f.projected}">£${formatCurrency(f.projected)}</td>
            <td data-value="${f.growth}" class="${f.growth >= 0 ? 'positive-growth' : 'negative-growth'}">
                £${formatCurrency(Math.abs(f.growth))}
                ${f.growth >= 0 ? '↑' : '↓'}
            </td>
        `;
        tbody.appendChild(row);
    });

    // Update forecast chart
    updateForecastChart(forecasts);
}

// Update forecast chart
function updateForecastChart(forecasts) {
    const current = getCurrentSnapshot();
    const currentTotal = current.reduce((sum, acc) =>
        sum + convertToGBP(acc.Balance, acc.Currency), 0);

    // Generate monthly data points
    const labels = [];
    const data = [];
    const monthlyContribution = parseFloat(document.getElementById('monthlyContribution').value) || 0;

    for (let month = 0; month <= forecastMonths; month++) {
        labels.push(month === 0 ? 'Now' : `+${month}mo`);

        if (month === 0) {
            data.push(currentTotal);
        } else {
            // Recalculate for this specific month
            let totalForMonth = 0;
            current.forEach(acc => {
                const principal = convertToGBP(acc.Balance, acc.Currency);
                const annualRate = acc.Interest_Rate / 100;
                const monthlyRate = annualRate / 12;

                let futureValue = principal * Math.pow(1 + monthlyRate, month);

                if (monthlyContribution > 0) {
                    const proportion = principal / currentTotal;
                    const accountContribution = monthlyContribution * proportion;

                    if (monthlyRate > 0) {
                        futureValue += accountContribution *
                            (Math.pow(1 + monthlyRate, month) - 1) / monthlyRate;
                    } else {
                        futureValue += accountContribution * month;
                    }
                }

                totalForMonth += futureValue;
            });
            data.push(totalForMonth);
        }
    }

    const ctx = document.getElementById('forecastChart').getContext('2d');
    if (charts.forecast) {
        charts.forecast.destroy();
    }

    charts.forecast = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Projected Net Worth',
                data: data,
                borderColor: '#3b82f6',
                backgroundColor: 'rgba(59, 130, 246, 0.1)',
                borderWidth: 3,
                borderDash: [5, 5],
                tension: 0.4,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            animation: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return `£${formatCurrency(context.parsed.y)}`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: false,
                    ticks: {
                        callback: function(value) {
                            return '£' + (value / 1000).toFixed(0) + 'k';
                        }
                    }
                }
            }
        }
    });
}

// Setup net worth event listeners
function setupNetWorthEventListeners() {
    // Forecast preset buttons
    document.querySelectorAll('.forecast-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const months = parseInt(e.target.dataset.months);
            forecastMonths = months;

            // Update active state
            document.querySelectorAll('.forecast-btn').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');

            // Update forecast date input
            const newDate = new Date();
            newDate.setMonth(newDate.getMonth() + months);
            document.getElementById('forecastDate').valueAsDate = newDate;

            updateForecast();
        });
    });

    // Forecast date picker
    document.getElementById('forecastDate').addEventListener('change', (e) => {
        const selectedDate = new Date(e.target.value);
        const now = new Date();
        const monthsDiff = Math.round((selectedDate - now) / (1000 * 60 * 60 * 24 * 30.44));
        forecastMonths = Math.max(1, monthsDiff);

        // Clear preset button active states
        document.querySelectorAll('.forecast-btn').forEach(b => b.classList.remove('active'));

        updateForecast();
    });

    // Monthly contribution input
    document.getElementById('monthlyContribution').addEventListener('input', () => {
        updateForecast();
    });

    // Make tables sortable
    makeSortable('holdingsTable');
    makeSortable('forecastTable');
}
