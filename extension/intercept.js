/**
 * SharePoint PDF Opener — Intercept Page Script
 * 
 * This page is loaded when declarativeNetRequest redirects a direct
 * .pdf URL navigation on SharePoint. It retrieves the original URL
 * from the service worker and triggers the download+open pipeline.
 */

(function () {
  'use strict';

  const filenameEl = document.getElementById('filename');
  const statusEl = document.getElementById('status');
  const errorEl = document.getElementById('error');

  /**
   * Main flow:
   * 1. Ask the service worker for the original intercepted URL
   * 2. Tell the service worker to download and open it
   * 3. Close this tab (or navigate back)
   */
  async function init() {
    try {
      // Step 1: Get the original URL from the service worker
      const response = await chrome.runtime.sendMessage({ type: 'GET_PENDING_URL' });

      if (!response || !response.url) {
        showError('Could not determine the PDF URL. The page may have already been processed.');
        scheduleClose(3000);
        return;
      }

      const pdfUrl = response.url;
      const filename = extractFilename(pdfUrl);

      filenameEl.textContent = filename;
      statusEl.textContent = 'Downloading from SharePoint…';

      // Step 2: Append ?download=1 and send to the service worker
      const downloadUrl = ensureDownloadParam(pdfUrl);

      const openResponse = await chrome.runtime.sendMessage({
        type: 'OPEN_PDF',
        url: downloadUrl,
        filename: filename
      });

      if (openResponse?.status === 'ok') {
        statusEl.textContent = 'Download started — opening in Acrobat…';
        // Close this tab after a brief delay to let the user see the status
        scheduleClose(2000);
      } else {
        showError(openResponse?.message || 'Failed to start download.');
        scheduleClose(5000);
      }

    } catch (err) {
      showError(`Error: ${err.message}`);
      scheduleClose(5000);
    }
  }

  function ensureDownloadParam(url) {
    try {
      const u = new URL(url);
      u.searchParams.set('download', '1');
      return u.toString();
    } catch {
      return url;
    }
  }

  function extractFilename(url) {
    try {
      const u = new URL(url);
      const parts = u.pathname.split('/');
      for (let i = parts.length - 1; i >= 0; i--) {
        const part = decodeURIComponent(parts[i]);
        if (part.toLowerCase().endsWith('.pdf')) return part;
      }
    } catch { /* fall through */ }
    return 'document.pdf';
  }

  function showError(message) {
    errorEl.textContent = message;
    errorEl.style.display = 'block';
    statusEl.textContent = '';
  }

  function scheduleClose(delayMs) {
    setTimeout(() => {
      // Try to go back first; if there's no history, close the tab
      if (window.history.length > 1) {
        window.history.back();
      } else {
        window.close();
      }
    }, delayMs);
  }

  // Run on page load
  init();
})();
