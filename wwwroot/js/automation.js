// Automation and SignalR integration

let connection = null;
let isRunning = false;

/**
 * Initialize SignalR connection
 */
async function initSignalR() {
    connection = new signalR.HubConnectionBuilder()
        .withUrl("/automationHub")
        .withAutomaticReconnect()
        .build();

    // Handle log messages
    connection.on("ReceiveLog", function (logEntry) {
        addLogEntry(logEntry);
    });

    // Handle progress updates
    connection.on("ReceiveProgress", function (current, total, status) {
        updateProgress(current, total, status);
    });

    // Handle completion
    connection.on("ReceiveComplete", function (result) {
        handleComplete(result);
    });

    try {
        await connection.start();
        console.log("SignalR connected");
    } catch (err) {
        console.error("SignalR connection error:", err);
        // Retry after delay
        setTimeout(initSignalR, 5000);
    }
}

/**
 * Add a log entry to the console
 * @param {object} logEntry - The log entry
 */
function addLogEntry(logEntry) {
    const logContainer = document.getElementById('log-container');
    if (!logContainer) return;

    const entry = document.createElement('div');
    entry.className = `log-entry log-${logEntry.level.toLowerCase()}`;

    const timestamp = new Date(logEntry.timestamp).toLocaleTimeString();
    const icon = getLogIcon(logEntry.level);

    entry.innerHTML = `
        <span class="log-time">${timestamp}</span>
        <span class="log-icon">${icon}</span>
        <span class="log-message">${escapeHtml(logEntry.message)}</span>
        ${logEntry.details ? `<pre class="log-details">${escapeHtml(logEntry.details)}</pre>` : ''}
    `;

    logContainer.appendChild(entry);
    logContainer.scrollTop = logContainer.scrollHeight;
}

/**
 * Get icon for log level
 * @param {string} level - Log level
 * @returns {string} Icon HTML
 */
function getLogIcon(level) {
    switch (level.toLowerCase()) {
        case 'success': return '✓';
        case 'error': return '✗';
        case 'warning': return '⚠';
        default: return 'ℹ';
    }
}

/**
 * Update progress bar
 * @param {number} current - Current item
 * @param {number} total - Total items
 * @param {string} status - Status message
 */
function updateProgress(current, total, status) {
    const progressBar = document.getElementById('progress-bar');
    const progressText = document.getElementById('progress-text');

    if (progressBar) {
        const percent = total > 0 ? (current / total) * 100 : 0;
        progressBar.style.width = percent + '%';
        progressBar.setAttribute('aria-valuenow', percent);
    }

    if (progressText) {
        progressText.textContent = `${current}/${total} - ${status}`;
    }
}

/**
 * Handle automation completion
 * @param {object} result - Automation result
 */
function handleComplete(result) {
    isRunning = false;

    // Update UI
    const confirmBtn = document.getElementById('btn-confirm');
    if (confirmBtn) {
        confirmBtn.disabled = false;
        confirmBtn.innerHTML = '<i class="bi bi-play-fill"></i> Run Automation';
    }

    // Show summary
    const summaryEl = document.getElementById('automation-summary');
    if (summaryEl) {
        summaryEl.style.display = 'block';
        summaryEl.innerHTML = `
            <div class="alert ${result.success ? 'alert-success' : 'alert-warning'}">
                <h5>${result.success ? 'Automation Completed Successfully' : 'Automation Completed with Issues'}</h5>
                <p>Processed: ${result.successfulItems}/${result.totalItems} items</p>
                ${result.failedItems > 0 ? `<p class="text-danger">Failed: ${result.failedItems} items</p>` : ''}
                ${result.screenshotPath ? `<a href="/Home/DownloadFile?path=${encodeURIComponent(result.screenshotPath)}" class="btn btn-sm btn-outline-primary me-2">Download Screenshot</a>` : ''}
                ${result.tracePath ? `<a href="/Home/DownloadFile?path=${encodeURIComponent(result.tracePath)}" class="btn btn-sm btn-outline-primary">Download Trace</a>` : ''}
            </div>
        `;
    }
}

/**
 * Start the automation
 */
async function startAutomation() {
    if (isRunning) {
        console.log("Automation already running");
        return;
    }

    isRunning = true;

    // Clear previous logs
    const logContainer = document.getElementById('log-container');
    if (logContainer) {
        logContainer.innerHTML = '';
    }

    // Hide previous summary
    const summaryEl = document.getElementById('automation-summary');
    if (summaryEl) {
        summaryEl.style.display = 'none';
    }

    // Update button state
    const confirmBtn = document.getElementById('btn-confirm');
    if (confirmBtn) {
        confirmBtn.disabled = true;
        confirmBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Running...';
    }

    // Reset progress
    updateProgress(0, 0, 'Starting...');

    // Get form data
    const form = document.getElementById('automation-form');
    const formData = new FormData(form);

    // Get selected profiles
    const selectedProfiles = [];
    document.querySelectorAll('.profile-checkbox:checked').forEach(cb => {
        selectedProfiles.push(parseInt(cb.value));
    });
    formData.append('selectedProfileIds', JSON.stringify(selectedProfiles));

    try {
        const response = await fetch('/Home/RunAutomation', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (!result.success && !isRunning) {
            // Handle immediate failure
            handleComplete(result);
        }
    } catch (error) {
        console.error('Error starting automation:', error);
        addLogEntry({
            timestamp: new Date().toISOString(),
            level: 'Error',
            message: 'Failed to start automation: ' + error.message
        });
        isRunning = false;

        const confirmBtn = document.getElementById('btn-confirm');
        if (confirmBtn) {
            confirmBtn.disabled = false;
            confirmBtn.innerHTML = '<i class="bi bi-play-fill"></i> Run Automation';
        }
    }
}

/**
 * Escape HTML characters
 * @param {string} text - Text to escape
 * @returns {string} Escaped text
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * Select/deselect all profiles
 * @param {boolean} select - Whether to select all
 */
function selectAllProfiles(select) {
    document.querySelectorAll('.profile-checkbox').forEach(cb => {
        cb.checked = select;
    });
    updateSelectedCount();
}

/**
 * Update the selected count display
 */
function updateSelectedCount() {
    const count = document.querySelectorAll('.profile-checkbox:checked').length;
    const total = document.querySelectorAll('.profile-checkbox').length;
    const countEl = document.getElementById('selected-count');
    if (countEl) {
        countEl.textContent = `${count}/${total} selected`;
    }
}

// Initialize on page load
document.addEventListener('DOMContentLoaded', function () {
    initSignalR();

    // Add change handler to checkboxes
    document.querySelectorAll('.profile-checkbox').forEach(cb => {
        cb.addEventListener('change', updateSelectedCount);
    });

    updateSelectedCount();
});

// Export functions
window.automation = {
    start: startAutomation,
    selectAll: selectAllProfiles
};
