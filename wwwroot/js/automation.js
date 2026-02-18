// Automation and SignalR integration with real-time progress dashboard

let connection = null;
let isRunning = false;
let filledCount = 0;
let unfilledCount = 0;
let totalItemCount = 0;
let logText = '';
let currentItemType = '';
let filledItemsIndex = 0;

async function initSignalR() {
    connection = new signalR.HubConnectionBuilder()
        .withUrl("/automationHub")
        .withAutomaticReconnect()
        .build();

    connection.on("ReceiveLog", function (logEntry) {
        addLogEntry(logEntry);
        parseLogForProgress(logEntry);
    });

    connection.on("ReceiveProgress", function (current, total, status) {
        updateProgress(current, total, status);
    });

    connection.on("ReceiveComplete", function (result) {
        handleComplete(result);
    });

    try {
        await connection.start();
        console.log("SignalR connected");
    } catch (err) {
        console.error("SignalR connection error:", err);
        setTimeout(initSignalR, 5000);
    }
}

function normalizeLevel(level) {
    if (typeof level === 'number') {
        return ['info', 'success', 'warning', 'error'][level] || 'info';
    }
    return (level || 'info').toString().toLowerCase();
}

function addLogEntry(logEntry) {
    const logContainer = document.getElementById('log-container');
    if (!logContainer) return;

    const level = normalizeLevel(logEntry.level);
    const entry = document.createElement('div');
    entry.className = `log-entry log-${level}`;

    const timestamp = new Date(logEntry.timestamp).toLocaleTimeString();
    const icon = getLogIcon(level);

    entry.innerHTML = `
        <span class="log-time">${timestamp}</span>
        <span class="log-icon">${icon}</span>
        <span class="log-message">${escapeHtml(logEntry.message)}</span>
        ${logEntry.details ? `<pre class="log-details">${escapeHtml(logEntry.details)}</pre>` : ''}
    `;

    logContainer.appendChild(entry);
    logContainer.scrollTop = logContainer.scrollHeight;

    // Accumulate log text for download
    logText += `[${timestamp}] [${level.toUpperCase().padEnd(7)}] ${logEntry.message}\n`;
    if (logEntry.details) logText += `    Details: ${logEntry.details}\n`;
}

function getLogIcon(level) {
    switch (level) {
        case 'success': return '✓';
        case 'error': return '✗';
        case 'warning': return '⚠';
        default: return 'i';
    }
}

function parseLogForProgress(logEntry) {
    const msg = logEntry.message || '';
    const level = normalizeLevel(logEntry.level);

    // Step: Login
    if (msg.includes('Navigating to Cortizo Center login')) {
        setActiveStep('login');
        setProgressLabel('Logging into Cortizo Center...');
        setProgressBar(5);
    }
    else if (msg.includes('Login successful')) {
        completeStep('login');
        setProgressBar(15);
    }

    // Step: Header
    else if (msg.includes('Creating new valuation')) {
        setActiveStep('header');
        setProgressLabel('Creating new valuation...');
        setProgressBar(18);
    }
    else if (msg.includes('Setting header fields')) {
        setProgressLabel('Setting quotation header fields...');
        setProgressBar(22);
    }
    else if (msg.includes('Setting customized prices')) {
        setProgressLabel('Setting customized prices...');
        setProgressBar(28);
    }
    else if (msg.includes('Customized prices set complete')) {
        completeStep('header');
        setProgressBar(30);
    }

    // Step: Profiles - start
    else if (msg.includes('Ensuring') && msg.includes('rows are available')) {
        setActiveStep('profiles');
        setProgressLabel('Preparing profile grid rows...');
        setProgressBar(32);
        showFilledSummary();
        showFilledItemsPanel();
        const m = msg.match(/(\d+) rows/);
        if (m) totalItemCount = parseInt(m[1]);
    }
    else if (msg.includes('Filling') && msg.includes('profile items')) {
        setActiveStep('profiles');
        currentItemType = 'Profile';
        const m = msg.match(/Filling (\d+)/);
        if (m) totalItemCount = parseInt(m[1]);
        document.getElementById('stat-total-items').textContent = totalItemCount;
        setProgressLabel(`Filling ${totalItemCount} profiles...`);
        setProgressBar(35);
        showCurrentItem();
    }

    // Track individual profile rows: [ROW 0001] Processing: REF=2015, AMT=34, DESC=...
    else if (msg.includes('[ROW ') && msg.includes('Processing:')) {
        const refMatch = msg.match(/REF=(\S+?)(?:,|$)/);
        const amtMatch = msg.match(/AMT=(\d+)/);
        const descMatch = msg.match(/DESC=(.+)$/);
        const rowMatch = msg.match(/\[ROW\s+(\d+)\]/);

        showCurrentItem();
        document.getElementById('current-item-title').textContent = 'Filling Profile...';
        if (refMatch) document.getElementById('current-item-ref').textContent = refMatch[1];
        if (amtMatch) document.getElementById('current-item-qty').textContent = amtMatch[1];
        if (descMatch) document.getElementById('current-item-desc').textContent = descMatch[1].trim();
        if (rowMatch) {
            const rowNum = parseInt(rowMatch[1]);
            document.getElementById('current-item-counter').textContent = `${rowNum}/${totalItemCount}`;
            const pct = 35 + ((rowNum / Math.max(totalItemCount, 1)) * 30);
            setProgressBar(Math.min(pct, 65));
            setProgressLabel(`Filling profile ${rowNum}/${totalItemCount}...`);
        }
        // Reset finish/shade display for this row
        document.getElementById('current-item-finish').textContent = '-';
        document.getElementById('current-item-shade').textContent = '-';
    }

    // Profile row amount calculated
    else if (msg.includes('[ROW ') && msg.includes('Amount calculated')) {
        const amountMatch = msg.match(/:\s*(.+?)$/);
        const rowMatch = msg.match(/\[ROW\s+(\d+)\]/);
        filledCount++;
        document.getElementById('stat-filled').textContent = filledCount;
        addFilledItemRow('Profile', amountMatch ? amountMatch[1].trim() : 'OK', true, rowMatch);
    }
    // Profile row amount NOT calculated
    else if (msg.includes('[ROW ') && msg.includes('Amount still not calculated')) {
        const rowMatch = msg.match(/\[ROW\s+(\d+)\]/);
        unfilledCount++;
        document.getElementById('stat-unfilled').textContent = unfilledCount;
        addFilledItemRow('Profile', '-', false, rowMatch);
    }

    // Profile batch complete
    else if (msg.includes('profile rows filled successfully')) {
        completeStep('profiles');
        setProgressBar(65);
        hideCurrentItem();
    }
    else if (msg.includes('Profile batch fill failed') && !msg.includes('canceled')) {
        const s = document.getElementById('step-profiles');
        if (s) { s.classList.remove('active'); s.classList.add('error'); }
        setProgressBar(65);
        hideCurrentItem();
    }

    // Step: Accessories start
    else if (msg.includes('Filling') && msg.includes('accessories')) {
        setActiveStep('accessories');
        currentItemType = 'Acc';
        const m = msg.match(/Filling (\d+)/);
        let accCount = 0;
        if (m) accCount = parseInt(m[1]);
        totalItemCount += accCount;
        document.getElementById('stat-total-items').textContent = totalItemCount;
        setProgressLabel(`Filling ${accCount} accessories...`);
        setProgressBar(70);
        showCurrentItem();
        document.getElementById('current-item-title').textContent = 'Filling Accessory...';
    }

    // Track accessory rows: [ACC 0001] Processing: REF=222085, AMT=104, DESC=Die cut cleat
    else if (msg.includes('[ACC ') && msg.includes('Processing:')) {
        const refMatch = msg.match(/REF=(\S+?)(?:,|$)/);
        const amtMatch = msg.match(/AMT=(\d+)/);
        const descMatch = msg.match(/DESC=(.+)$/);
        const rowMatch = msg.match(/\[ACC\s+(\d+)\]/);

        showCurrentItem();
        document.getElementById('current-item-title').textContent = 'Filling Accessory...';
        if (refMatch) document.getElementById('current-item-ref').textContent = refMatch[1];
        if (amtMatch) document.getElementById('current-item-qty').textContent = amtMatch[1];
        if (descMatch) document.getElementById('current-item-desc').textContent = descMatch[1].trim();
        if (rowMatch) {
            document.getElementById('current-item-counter').textContent = `Acc ${rowMatch[1]}`;
        }
        document.getElementById('current-item-finish').textContent = '-';
        document.getElementById('current-item-shade').textContent = '-';
    }

    // Accessory amount calculated
    else if (msg.includes('[ACC ') && msg.includes('Amount calculated')) {
        const amountMatch = msg.match(/Amount calculated[^:]*:\s*([^,]+)/);
        const rowMatch = msg.match(/\[ACC\s+(\d+)\]/);
        filledCount++;
        document.getElementById('stat-filled').textContent = filledCount;
        addFilledItemRow('Acc', amountMatch ? amountMatch[1].trim() : 'OK', true, rowMatch);
    }
    // Accessory amount NOT calculated  
    else if (msg.includes('[ACC ') && (msg.includes('Amount still not calculated') || msg.includes('Amount not calculated'))) {
        const rowMatch = msg.match(/\[ACC\s+(\d+)\]/);
        // Only count the final failure, not the intermediate retry message
        if (msg.includes('still not calculated') || !msg.includes('retrying')) {
            unfilledCount++;
            document.getElementById('stat-unfilled').textContent = unfilledCount;
            addFilledItemRow('Acc', '-', false, rowMatch);
        }
    }

    else if (msg.includes('accessories filled successfully') || msg.includes('Accessory fill complete')) {
        completeStep('accessories');
        setProgressBar(85);
        hideCurrentItem();
    }

    // Step: Total
    else if (msg.includes('ESTIMATE TOTAL')) {
        setActiveStep('total');
        const totalMatch = msg.match(/([\d,.]+)\s*EUR/);
        if (totalMatch) {
            document.getElementById('stat-cortizo-total').textContent = totalMatch[1];
        }
        setProgressLabel('Extracting total...');
        setProgressBar(90);
    }

    // Unfilled warnings
    else if (msg.includes('UNFILLED PROFILES') || msg.includes('UNFILLED ACCESSORIES')) {
        // Already counted via individual rows - no double-counting
    }

    // Generating report
    else if (msg.includes('Generating report')) {
        setProgressLabel('Generating report...');
        setProgressBar(93);
    }

    // Creating proforma
    else if (msg.includes('Creating proforma')) {
        setProgressLabel('Creating proforma...');
        setProgressBar(95);
    }

    // Stopped by user
    else if (msg.includes('stopped by user') || msg.includes('Automation was stopped')) {
        setStatusBadge('Stopped', 'warning');
        setProgressLabel('Automation stopped by user');
        hideCurrentItem();
    }

    // Error
    else if (level === 'error' && msg.includes('failed') && !msg.includes('[ROW') && !msg.includes('[ACC')) {
        const activeStep = document.querySelector('.step-item.active');
        if (activeStep) {
            activeStep.classList.remove('active');
            activeStep.classList.add('error');
        }
        setStatusBadge('Error', 'danger');
    }

    // Completed
    else if (msg.includes('Automation completed')) {
        completeStep('total');
        setProgressBar(100);
        setProgressLabel('Automation completed!');
    }
}

// Track last processing row info for filled items table
let lastProcessingInfo = { ref: '', amt: '', desc: '', type: '' };

function addFilledItemRow(type, amount, success, rowMatch) {
    const tbody = document.getElementById('filled-items-tbody');
    if (!tbody) return;

    filledItemsIndex++;

    const ref = document.getElementById('current-item-ref')?.textContent || '-';
    const qty = document.getElementById('current-item-qty')?.textContent || '-';
    const desc = document.getElementById('current-item-desc')?.textContent || '-';

    const statusBadge = success
        ? '<span class="badge bg-success">OK</span>'
        : '<span class="badge bg-warning text-dark">Missing</span>';

    const rowClass = success ? '' : 'table-warning';

    const tr = document.createElement('tr');
    tr.className = rowClass;
    tr.innerHTML = `
        <td>${filledItemsIndex}</td>
        <td><span class="badge ${type === 'Profile' ? 'bg-primary' : 'bg-info text-dark'}">${type}</span></td>
        <td class="fw-bold">${escapeHtml(ref)}</td>
        <td>${escapeHtml(qty)}</td>
        <td class="text-truncate" style="max-width: 180px;" title="${escapeHtml(desc)}">${escapeHtml(desc)}</td>
        <td class="fw-bold ${success ? 'text-success' : 'text-danger'}">${escapeHtml(amount)}</td>
        <td>${statusBadge}</td>
    `;
    tbody.appendChild(tr);

    // Update count badge
    const countEl = document.getElementById('filled-items-count');
    if (countEl) countEl.textContent = filledItemsIndex;

    // Scroll table to bottom
    const tableContainer = tbody.closest('div[style*="overflow"]');
    if (tableContainer) tableContainer.scrollTop = tableContainer.scrollHeight;
}

function showFilledItemsPanel() {
    const panel = document.getElementById('filled-items-panel');
    if (panel) panel.style.display = 'block';
}

function setActiveStep(stepId) {
    const step = document.getElementById('step-' + stepId);
    if (!step || step.classList.contains('completed')) return;

    document.querySelectorAll('.step-item.active').forEach(el => {
        if (el.id !== 'step-' + stepId) {
            el.classList.remove('active');
            el.classList.add('completed');
        }
    });
    step.classList.add('active');
    setStatusBadge('Running', 'primary');
}

function completeStep(stepId) {
    const step = document.getElementById('step-' + stepId);
    if (step) {
        step.classList.remove('active');
        step.classList.add('completed');
    }
}

function setProgressBar(percent) {
    const bar = document.getElementById('progress-bar');
    if (bar) {
        bar.style.width = percent + '%';
        if (percent >= 100) {
            bar.classList.remove('progress-bar-striped');
        } else {
            bar.classList.add('progress-bar-striped');
        }
    }
    const label = document.getElementById('progress-percent-label');
    if (label) label.textContent = Math.round(percent) + '%';
}

function setProgressLabel(text) {
    const el = document.getElementById('progress-step-label');
    if (el) el.textContent = text;
    const sub = document.getElementById('progress-text');
    if (sub) sub.textContent = text;
}

function setStatusBadge(text, color) {
    const badge = document.getElementById('automation-status-badge');
    if (badge) {
        badge.textContent = text;
        badge.className = `badge bg-${color}`;
    }
}

function showCurrentItem() {
    const el = document.getElementById('current-item-detail');
    if (el) el.style.display = 'block';
}

function hideCurrentItem() {
    const el = document.getElementById('current-item-detail');
    if (el) el.style.display = 'none';
}

function showFilledSummary() {
    const el = document.getElementById('filled-items-summary');
    if (el) el.style.display = 'block';
}

function updateProgress(current, total, status) {
    const percent = total > 0 ? (current / total) * 100 : 0;
    setProgressBar(percent);
    setProgressLabel(`${current}/${total} - ${status}`);
}

function handleComplete(result) {
    isRunning = false;

    hideStopButton();
    const confirmBtn = document.getElementById('btn-confirm');
    if (confirmBtn) {
        confirmBtn.disabled = false;
        confirmBtn.innerHTML = '<i class="bi bi-play-fill me-2"></i>Run Automation';
    }

    hideCurrentItem();
    showFilledSummary();

    // Show download log button
    const dlBtn = document.getElementById('btn-download-log');
    if (dlBtn) dlBtn.style.display = 'inline-block';

    document.getElementById('stat-filled').textContent = result.successfulItems || filledCount;
    document.getElementById('stat-unfilled').textContent =
        (result.unfilledProfiles ? result.unfilledProfiles.length : 0) +
        (result.unfilledAccessories ? result.unfilledAccessories.length : 0);
    document.getElementById('stat-total-items').textContent = result.totalItems || totalItemCount;

    if (result.cortizoTotal > 0) {
        document.getElementById('stat-cortizo-total').textContent = result.cortizoTotal.toFixed(2);
    }

    if (result.success) {
        setStatusBadge('Completed', 'success');
        setProgressBar(100);
        setProgressLabel('Automation completed successfully!');
        completeStep('total');
    } else {
        setStatusBadge('Completed (issues)', 'warning');
        setProgressBar(100);
        setProgressLabel('Completed with issues - review unfilled items');
    }

    const summaryEl = document.getElementById('automation-summary');
    if (summaryEl) {
        summaryEl.style.display = 'block';
        let unfilledHtml = '';

        if (result.unfilledProfiles && result.unfilledProfiles.length > 0) {
            unfilledHtml += '<div class="mt-2"><strong>Unfilled Profiles:</strong><ul class="mb-1">';
            result.unfilledProfiles.forEach(item => {
                unfilledHtml += `<li>Row ${item.rowNumber}: REF ${item.refNumber} x ${item.amount} - ${item.reason}</li>`;
            });
            unfilledHtml += '</ul></div>';
        }

        if (result.unfilledAccessories && result.unfilledAccessories.length > 0) {
            unfilledHtml += '<div class="mt-2"><strong>Unfilled Accessories:</strong><ul class="mb-1">';
            result.unfilledAccessories.forEach(item => {
                unfilledHtml += `<li>Row ${item.rowNumber}: REF ${item.refNumber} x ${item.amount} - ${item.reason}</li>`;
            });
            unfilledHtml += '</ul></div>';
        }

        summaryEl.innerHTML = `
            <div class="alert ${result.success ? 'alert-success' : 'alert-warning'} mb-0">
                <h6 class="alert-heading mb-1">${result.success ? 'Automation Completed' : 'Completed with Issues'}</h6>
                <small>Processed: ${result.successfulItems}/${result.totalItems} items
                ${result.cortizoTotal > 0 ? ` | Cortizo Total: <strong>${result.cortizoTotal.toFixed(2)} EUR</strong>` : ''}</small>
                ${unfilledHtml}
                <div class="mt-2">
                    ${result.screenshotPath ? `<a href="/Home/DownloadFile?path=${encodeURIComponent(result.screenshotPath)}" class="btn btn-sm btn-outline-primary me-1">Screenshot</a>` : ''}
                    ${result.tracePath ? `<a href="/Home/DownloadFile?path=${encodeURIComponent(result.tracePath)}" class="btn btn-sm btn-outline-primary me-1">Server Log</a>` : ''}
                    <button class="btn btn-sm btn-outline-secondary" onclick="automation.downloadLog()"><i class="bi bi-download me-1"></i>Download UI Log</button>
                </div>
            </div>
        `;
    }
}

function downloadLog() {
    if (!logText) return;
    const blob = new Blob([logText], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `automation-log-${new Date().toISOString().slice(0,19).replace(/[T:]/g,'-')}.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

async function stopAutomation() {
    if (!isRunning) return;

    const stopBtn = document.getElementById('btn-stop');
    if (stopBtn) {
        stopBtn.disabled = true;
        stopBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Stopping...';
    }

    try {
        await fetch('/Home/StopAutomation', { method: 'POST' });
        addLogEntry({
            timestamp: new Date().toISOString(),
            level: 'Warning',
            message: 'Stop requested - waiting for current operation to finish...'
        });
        setStatusBadge('Stopping', 'warning');
    } catch (error) {
        console.error('Failed to stop automation:', error);
    }
}

function showStopButton() {
    const stopBtn = document.getElementById('btn-stop');
    if (stopBtn) {
        stopBtn.style.display = 'block';
        stopBtn.disabled = false;
        stopBtn.innerHTML = '<i class="bi bi-stop-fill me-1"></i>Stop';
    }
}

function hideStopButton() {
    const stopBtn = document.getElementById('btn-stop');
    if (stopBtn) {
        stopBtn.style.display = 'none';
    }
}

async function startAutomation() {
    if (isRunning) return;

    isRunning = true;
    filledCount = 0;
    unfilledCount = 0;
    totalItemCount = 0;
    filledItemsIndex = 0;
    logText = '';

    // Reset dashboard
    document.getElementById('log-container').innerHTML = '';
    document.querySelectorAll('.step-item').forEach(el => {
        el.classList.remove('active', 'completed', 'error');
    });
    setProgressBar(0);
    setProgressLabel('Starting automation...');
    setStatusBadge('Starting', 'info');
    document.getElementById('stat-filled').textContent = '0';
    document.getElementById('stat-unfilled').textContent = '0';
    document.getElementById('stat-total-items').textContent = '0';
    document.getElementById('stat-cortizo-total').textContent = '-';
    hideCurrentItem();
    showFilledSummary();

    // Reset filled items table
    const tbody = document.getElementById('filled-items-tbody');
    if (tbody) tbody.innerHTML = '';
    const countEl = document.getElementById('filled-items-count');
    if (countEl) countEl.textContent = '0';

    // Hide download button
    const dlBtn = document.getElementById('btn-download-log');
    if (dlBtn) dlBtn.style.display = 'none';

    const summaryEl = document.getElementById('automation-summary');
    if (summaryEl) summaryEl.style.display = 'none';

    const confirmBtn = document.getElementById('btn-confirm');
    if (confirmBtn) {
        confirmBtn.disabled = true;
        confirmBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Running...';
    }

    showStopButton();

    const form = document.getElementById('automation-form');
    const formData = new FormData(form);

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
        setStatusBadge('Error', 'danger');

        if (confirmBtn) {
            confirmBtn.disabled = false;
            confirmBtn.innerHTML = '<i class="bi bi-play-fill me-2"></i>Run Automation';
        }
    }
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function selectAllProfiles(select) {
    document.querySelectorAll('.profile-checkbox').forEach(cb => {
        cb.checked = select;
    });
    updateSelectedCount();
}

function updateSelectedCount() {
    const count = document.querySelectorAll('.profile-checkbox:checked').length;
    const total = document.querySelectorAll('.profile-checkbox').length;
    const countEl = document.getElementById('selected-count');
    if (countEl) countEl.textContent = `${count}/${total} selected`;
}

document.addEventListener('DOMContentLoaded', function () {
    initSignalR();
    document.querySelectorAll('.profile-checkbox').forEach(cb => {
        cb.addEventListener('change', updateSelectedCount);
    });
    updateSelectedCount();
});

window.automation = {
    start: startAutomation,
    stop: stopAutomation,
    selectAll: selectAllProfiles,
    downloadLog: downloadLog
};
