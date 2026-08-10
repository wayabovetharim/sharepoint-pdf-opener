/**
 * SharePoint PDF Opener — Background Service Worker
 * 
 * Orchestrates the entire PDF interception pipeline:
 *   1. Receives PDF URLs from the content script or intercept page
 *   2. Downloads the PDF using the browser's authenticated session
 *   3. Opens the downloaded file via chrome.downloads.open() (uses OS file associations)
 *   4. Manages DNR rules for direct URL navigation interception
 *   5. Handles enable/disable toggle and cleanup tasks
 *
 * NOTE: No native messaging required. Files open via Windows file
 * associations — if Acrobat is the default PDF handler, it opens in Acrobat.
 * Alternatively, user can set Edge's "Always open files of this type" for PDFs.
 */

'use strict';

// ── Constants ────────────────────────────────────────────────────────────
const DOWNLOAD_SUBDIR = 'SharePointPDFs';
const CLEANUP_MAX_AGE_HOURS = 24;
const DNR_RULE_ID_PDF_REDIRECT = 1;

// ── State ────────────────────────────────────────────────────────────────

/**
 * Map of tabId → original URL, used to pass the intercepted URL
 * from webNavigation to the intercept.html page.
 */
const pendingInterceptions = new Map();

/**
 * Map of downloadId → metadata, to track in-flight downloads.
 */
const activeDownloads = new Map();

// ── Initialization ───────────────────────────────────────────────────────

chrome.runtime.onInstalled.addListener(async () => {
  // Set default enabled state
  const { enabled } = await chrome.storage.local.get({ enabled: true });
  await chrome.storage.local.set({ enabled, pdfCount: 0 });

  // Set up declarativeNetRequest rules for direct PDF URL interception
  await updateDnrRules(enabled);

  console.log('[SP PDF Opener] Extension installed. Enabled:', enabled);
});

chrome.runtime.onStartup.addListener(async () => {
  const { enabled } = await chrome.storage.local.get({ enabled: true });
  await updateDnrRules(enabled);
  console.log('[SP PDF Opener] Extension started. Enabled:', enabled);
});

// ── DNR Rule Management ──────────────────────────────────────────────────

/**
 * Adds or removes the declarativeNetRequest rule that redirects
 * direct .pdf URL navigations on SharePoint to our intercept page.
 */
async function updateDnrRules(enabled) {
  try {
    if (enabled) {
      await chrome.declarativeNetRequest.updateDynamicRules({
        removeRuleIds: [DNR_RULE_ID_PDF_REDIRECT],
        addRules: [{
          id: DNR_RULE_ID_PDF_REDIRECT,
          priority: 1,
          action: {
            type: 'redirect',
            redirect: { extensionPath: '/intercept.html' }
          },
          condition: {
            // Match direct .pdf URLs on SharePoint (main frame only)
            regexFilter: '^https://[^/]+\\.sharepoint\\.com/.+\\.pdf(\\?[^#]*)?(#.*)?$',
            resourceTypes: ['main_frame']
          }
        }]
      });
    } else {
      await chrome.declarativeNetRequest.updateDynamicRules({
        removeRuleIds: [DNR_RULE_ID_PDF_REDIRECT],
        addRules: []
      });
    }
  } catch (err) {
    console.error('[SP PDF Opener] DNR rule update failed:', err);
  }
}

// ── Web Navigation Monitoring ────────────────────────────────────────────

/**
 * Capture the original URL before DNR redirects it.
 * webNavigation.onBeforeNavigate fires before DNR takes effect,
 * so we can store the URL for the intercept page to retrieve.
 */
chrome.webNavigation.onBeforeNavigate.addListener(
  (details) => {
    if (details.frameId !== 0) return; // main frame only

    const url = details.url;
    if (isDirectPdfUrl(url)) {
      pendingInterceptions.set(details.tabId, {
        url: url,
        timestamp: Date.now()
      });

      // Clean up stale entries (older than 30 seconds)
      for (const [tabId, data] of pendingInterceptions) {
        if (Date.now() - data.timestamp > 30000) {
          pendingInterceptions.delete(tabId);
        }
      }
    }
  },
  { url: [{ hostSuffix: '.sharepoint.com' }] }
);

/**
 * Check if a URL is a direct PDF file on SharePoint.
 */
function isDirectPdfUrl(url) {
  try {
    const u = new URL(url);
    if (!u.hostname.endsWith('.sharepoint.com')) return false;
    return /\.pdf(\?[^#]*)?(#.*)?$/i.test(u.pathname);
  } catch {
    return false;
  }
}

// ── Message Handling ─────────────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  switch (msg.type) {

    case 'OPEN_PDF':
      handlePdfOpen(msg.url, msg.filename, sender.tab?.id)
        .then(() => sendResponse({ status: 'ok' }))
        .catch((err) => sendResponse({ status: 'error', message: err.message }));
      return true; // async response

    case 'GET_PENDING_URL':
      // Called by intercept.html to retrieve the original URL
      const tabId = sender.tab?.id;
      const pending = pendingInterceptions.get(tabId);
      if (pending) {
        pendingInterceptions.delete(tabId);
        sendResponse({ url: pending.url });
      } else {
        sendResponse({ url: null });
      }
      return false;

    case 'TOGGLE_ENABLED':
      handleToggle(msg.enabled);
      sendResponse({ status: 'ok' });
      return false;

    case 'OPEN_DOWNLOADS':
      openDownloadsFolder();
      sendResponse({ status: 'ok' });
      return false;

    default:
      return false;
  }
});

// ── PDF Download & Open Pipeline ─────────────────────────────────────────

/**
 * Main pipeline: download the PDF, then open it in Acrobat.
 */
async function handlePdfOpen(url, filename, tabId) {
  // Check if extension is enabled
  const { enabled } = await chrome.storage.local.get({ enabled: true });
  if (!enabled) return;

  // Sanitize filename
  const safeName = sanitizeFilename(filename || extractFilenameFromUrl(url));

  console.log(`[SP PDF Opener] Downloading: ${safeName} from ${url}`);

  try {
    // Trigger download using the browser's download manager
    // (automatically uses the browser's authenticated session / cookies)
    const downloadId = await new Promise((resolve, reject) => {
      chrome.downloads.download({
        url: url,
        filename: `${DOWNLOAD_SUBDIR}/${safeName}`,
        conflictAction: 'uniquify',
        saveAs: false
      }, (id) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else if (id === undefined) {
          reject(new Error('Download failed to start'));
        } else {
          resolve(id);
        }
      });
    });

    // Track this download
    activeDownloads.set(downloadId, {
      filename: safeName,
      tabId: tabId,
      startTime: Date.now()
    });

    // Navigate the tab back (if it was a content script interception)
    // Don't navigate back if it's the intercept page — that's handled by intercept.js
    if (tabId) {
      try {
        const tab = await chrome.tabs.get(tabId);
        // Only go back if we're still on the SharePoint page (content script case)
        if (tab.url?.includes('.sharepoint.com')) {
          // Don't navigate back — the user stays on the same page
          // They clicked a link and the PDF opens in Acrobat.
          // The click was intercepted, so no navigation occurred.
        }
      } catch { /* tab might have closed */ }
    }

  } catch (err) {
    console.error('[SP PDF Opener] Download failed:', err);
    throw err;
  }
}

// ── Download Completion Monitoring ───────────────────────────────────────

chrome.downloads.onChanged.addListener((delta) => {
  // Only care about our tracked downloads
  if (!activeDownloads.has(delta.id)) return;

  if (delta.state?.current === 'complete') {
    // Download finished — open the file using OS file associations
    const meta = activeDownloads.get(delta.id);
    activeDownloads.delete(delta.id);

    console.log(`[SP PDF Opener] Download complete: ${meta?.filename}`);
    openDownloadedFile(delta.id);
    incrementPdfCount();
  }

  if (delta.state?.current === 'interrupted') {
    const meta = activeDownloads.get(delta.id);
    activeDownloads.delete(delta.id);
    console.error(`[SP PDF Opener] Download interrupted: ${meta?.filename}`);
  }
});

// ── File Opening ─────────────────────────────────────────────────────────

/**
 * Opens a downloaded file using the OS default application.
 * If Acrobat is the default PDF handler (or Edge's "Always open files of
 * this type" is set for PDFs), the file opens in Acrobat automatically.
 *
 * No native messaging required — bypasses the NativeMessagingUserLevelHosts
 * corporate policy entirely.
 */
function openDownloadedFile(downloadId) {
  try {
    chrome.downloads.open(downloadId);
    console.log('[SP PDF Opener] File opened via OS file association');
  } catch (err) {
    console.error('[SP PDF Opener] Failed to open file:', err);
    // Fallback: show the file in the downloads folder so user can open manually
    chrome.downloads.show(downloadId);
  }
}

/**
 * Cleans up old downloaded PDFs using the downloads API directly.
 * No native messaging needed — uses chrome.downloads.search() and removeFile().
 */
async function cleanupOldDownloads() {
  try {
    const cutoffTime = new Date(Date.now() - CLEANUP_MAX_AGE_HOURS * 60 * 60 * 1000);
    const results = await chrome.downloads.search({
      filenameRegex: 'SharePointPDFs',
      endedBefore: cutoffTime.toISOString(),
      state: 'complete'
    });

    let deletedCount = 0;
    for (const item of results) {
      try {
        await chrome.downloads.removeFile(item.id);
        await chrome.downloads.erase({ id: item.id });
        deletedCount++;
      } catch {
        // File may already be removed or in use — skip
      }
    }

    if (deletedCount > 0) {
      console.log(`[SP PDF Opener] Cleanup: removed ${deletedCount} old file(s)`);
    }
  } catch {
    // Cleanup is non-critical
  }
}

// ── Enable/Disable Toggle ────────────────────────────────────────────────

async function handleToggle(enabled) {
  await chrome.storage.local.set({ enabled });
  await updateDnrRules(enabled);
  console.log(`[SP PDF Opener] ${enabled ? 'Enabled' : 'Disabled'}`);
}

// ── Utility Functions ────────────────────────────────────────────────────

function extractFilenameFromUrl(url) {
  try {
    const u = new URL(url);
    const pathParts = u.pathname.split('/');
    for (let i = pathParts.length - 1; i >= 0; i--) {
      const part = decodeURIComponent(pathParts[i]);
      if (part.toLowerCase().endsWith('.pdf')) return part;
    }

    // Check 'id' query parameter
    const idParam = u.searchParams.get('id');
    if (idParam) {
      const idParts = idParam.split('/');
      for (let i = idParts.length - 1; i >= 0; i--) {
        const part = decodeURIComponent(idParts[i]);
        if (part.toLowerCase().endsWith('.pdf')) return part;
      }
    }
  } catch { /* fall through */ }

  return `sharepoint_${Date.now()}.pdf`;
}

function sanitizeFilename(name) {
  // Remove path separators and invalid characters
  return name
    .replace(/[/\\:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    || `document_${Date.now()}.pdf`;
}

function incrementPdfCount() {
  chrome.storage.local.get({ pdfCount: 0 }, (result) => {
    chrome.storage.local.set({ pdfCount: result.pdfCount + 1 });
  });
}

function openDownloadsFolder() {
  // Get the default download directory and open the SharePointPDFs subfolder
  chrome.downloads.showDefaultFolder();
}

// ── Periodic Cleanup ─────────────────────────────────────────────────────

// Periodic cleanup using alarms API (no native messaging needed)
chrome.alarms?.create('cleanup', { periodInMinutes: 60 });

chrome.alarms?.onAlarm.addListener((alarm) => {
  if (alarm.name === 'cleanup') {
    cleanupOldDownloads();
  }
});
