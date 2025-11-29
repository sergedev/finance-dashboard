// Global variables to store data and charts
let transactions = [];
let categories = [];
let charts = {};
let selectedCategories = new Set();
let currentTimeFilter = 'all';

// Initialize the dashboard
document.addEventListener('DOMContentLoaded', () => {
    loadData();
    setupEventListeners();
    updateLastUpdatedTime();
});

// Load CSV data
async function loadData() {
    try {
        // Load transactions
        const transactionsResponse = await fetch('data/transactions.2025.csv');
        const transactionsText = await transactionsResponse.text();

        // Load categories
        const categoriesResponse = await fetch('data/categories.2025.csv');
        const categoriesText = await categoriesResponse.text();

        // Parse CSVs
        Papa.parse(transactionsText, {
            header: true,
            skipEmptyLines: true,
            complete: (results) => {
                transactions = results.data.map(row => ({
                    ...row,
                    Amount: parseAmount(row.Amount),
                    Balance: parseAmount(row.Balance),
                    Date: parseDate(row.Date)
                }));

                // Parse categories
                Papa.parse(categoriesText, {
                    header: true,
                    skipEmptyLines: true,
                    complete: (results) => {
                        categories = results.data;
                        initializeDashboard();
                    }
                });
            }
        });
    } catch (error) {
        console.error('Error loading data:', error);
        alert('Error loading data. Please ensure CSV files are in the data/ folder.');
    }
}

// Parse amount/balance values - handles £ symbols, commas, and whitespace
function parseAmount(amountStr) {
    if (!amountStr) return 0;
    // Remove £ symbol, commas, quotes, and whitespace, then parse
    const cleaned = amountStr.toString().replace(/[£,"\s]/g, '');
    return parseFloat(cleaned) || 0;
}

// Parse DD-MMM-YY date format
function parseDate(dateStr) {
    const months = {
        'Jan': 0, 'Feb': 1, 'Mar': 2, 'Apr': 3, 'May': 4, 'Jun': 5,
        'Jul': 6, 'Aug': 7, 'Sep': 8, 'Oct': 9, 'Nov': 10, 'Dec': 11
    };

    const parts = dateStr.split('-');
    const day = parseInt(parts[0]);
    const month = months[parts[1]];
    const year = 2000 + parseInt(parts[2]);

    return new Date(year, month, day);
}

// Initialize dashboard after data is loaded
function initializeDashboard() {
    // Get unique categories from categories.csv file
    const uniqueCategories = [...new Set(categories.map(c => c.Category))].filter(cat => cat);

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
            selectedCategories.clear();
            document.getElementById('toggleAll').textContent = 'Select All';
        } else {
            const uniqueCategories = [...new Set(transactions.map(t => t.Category))];
            uniqueCategories.forEach(cat => selectedCategories.add(cat));
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
