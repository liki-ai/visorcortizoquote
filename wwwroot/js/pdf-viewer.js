// PDF.js viewer wrapper
// Uses PDF.js from CDN (can be bundled later)

let pdfDoc = null;
let pageNum = 1;
let pageRendering = false;
let pageNumPending = null;
let scale = 1.2;
let canvas = null;
let ctx = null;

/**
 * Initialize the PDF viewer
 * @param {string} canvasId - The canvas element ID
 */
function initPdfViewer(canvasId) {
    canvas = document.getElementById(canvasId);
    if (canvas) {
        ctx = canvas.getContext('2d');
    }
}

/**
 * Load and display a PDF from a URL
 * @param {string} url - URL to the PDF file
 */
async function loadPdf(url) {
    try {
        // Show loading state
        showPdfLoading();

        // Load the PDF using PDF.js
        const loadingTask = pdfjsLib.getDocument(url);
        pdfDoc = await loadingTask.promise;

        // Update page count
        document.getElementById('page-count').textContent = pdfDoc.numPages;

        // Render first page
        pageNum = 1;
        await renderPage(pageNum);

        // Hide loading state
        hidePdfLoading();
    } catch (error) {
        console.error('Error loading PDF:', error);
        showPdfError('Failed to load PDF: ' + error.message);
    }
}

/**
 * Render a specific page
 * @param {number} num - Page number to render
 */
async function renderPage(num) {
    if (!pdfDoc || !canvas) return;

    pageRendering = true;

    try {
        // Get the page
        const page = await pdfDoc.getPage(num);

        // Calculate viewport
        const viewport = page.getViewport({ scale: scale });
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        // Render the page
        const renderContext = {
            canvasContext: ctx,
            viewport: viewport
        };

        await page.render(renderContext).promise;

        pageRendering = false;

        // Update page number display
        document.getElementById('page-num').textContent = num;

        // If there's a pending page, render it
        if (pageNumPending !== null) {
            renderPage(pageNumPending);
            pageNumPending = null;
        }
    } catch (error) {
        console.error('Error rendering page:', error);
        pageRendering = false;
    }
}

/**
 * Queue a page for rendering
 * @param {number} num - Page number to queue
 */
function queueRenderPage(num) {
    if (pageRendering) {
        pageNumPending = num;
    } else {
        renderPage(num);
    }
}

/**
 * Go to previous page
 */
function prevPage() {
    if (pageNum <= 1) return;
    pageNum--;
    queueRenderPage(pageNum);
}

/**
 * Go to next page
 */
function nextPage() {
    if (!pdfDoc || pageNum >= pdfDoc.numPages) return;
    pageNum++;
    queueRenderPage(pageNum);
}

/**
 * Zoom in
 */
function zoomIn() {
    scale += 0.2;
    if (scale > 3) scale = 3;
    queueRenderPage(pageNum);
}

/**
 * Zoom out
 */
function zoomOut() {
    scale -= 0.2;
    if (scale < 0.5) scale = 0.5;
    queueRenderPage(pageNum);
}

/**
 * Show loading state
 */
function showPdfLoading() {
    const container = document.getElementById('pdf-container');
    if (container) {
        container.classList.add('loading');
    }
}

/**
 * Hide loading state
 */
function hidePdfLoading() {
    const container = document.getElementById('pdf-container');
    if (container) {
        container.classList.remove('loading');
    }
}

/**
 * Show error message
 * @param {string} message - Error message to display
 */
function showPdfError(message) {
    const errorEl = document.getElementById('pdf-error');
    if (errorEl) {
        errorEl.textContent = message;
        errorEl.style.display = 'block';
    }
    hidePdfLoading();
}

// Export functions for global use
window.pdfViewer = {
    init: initPdfViewer,
    load: loadPdf,
    prevPage: prevPage,
    nextPage: nextPage,
    zoomIn: zoomIn,
    zoomOut: zoomOut
};
