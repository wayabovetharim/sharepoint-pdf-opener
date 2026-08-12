/**
 * SharePoint PDF Opener — Content Script
 * 
 * Injected into all *.sharepoint.com pages at document_start.
 * Intercepts clicks on PDF links and routes them through the extension's
 * download-and-open-in-Acrobat pipeline, bypassing the browser viewer.
 * 
 * Detection strategies:
 *   1. Click interception (capture phase) — catches <a> links to .pdf URLs
 *   2. SharePoint SPA navigation monitoring — catches programmatic navigations
 *   3. Viewer page detection — detects when the SP online PDF viewer loads
 */

(function () {
  'use strict';

  // ── Guard: only run once per page ──────────────────────────────────────
  if (window.__spPdfOpenerInjected) return;
  window.__spPdfOpenerInjected = true;

  console.log('[SP PDF Opener] Content script loaded on:', window.location.hostname);

  // ── State ──────────────────────────────────────────────────────────────
  let extensionEnabled = true;

  // Sync enabled state from storage
  chrome.storage.local.get({ enabled: true }, (result) => {
    extensionEnabled = result.enabled;
  });

  chrome.storage.onChanged.addListener((changes) => {
    if (changes.enabled) {
      extensionEnabled = changes.enabled.newValue;
    }
  });

  // ── URL Analysis Helpers ───────────────────────────────────────────────

  /**
   * Determines if a URL points to a PDF on SharePoint.
   * Handles: direct .pdf links, ?id= parameter links, AllItems.aspx with PDF id.
   */
  function isPdfUrl(url) {
    if (!url) return false;
    try {
      const u = new URL(url, window.location.origin);
      // Must be on a sharepoint.com domain
      if (!u.hostname.endsWith('.sharepoint.com')) return false;

      const pathname = decodeURIComponent(u.pathname).toLowerCase();

      // Direct .pdf URL
      if (pathname.endsWith('.pdf')) return true;

      // SharePoint viewer URL with file path in 'id' query parameter
      // e.g., /sites/SiteName/Shared Documents/Forms/AllItems.aspx?id=/sites/.../file.pdf
      const idParam = u.searchParams.get('id');
      if (idParam && decodeURIComponent(idParam).toLowerCase().endsWith('.pdf')) return true;

      // SharePoint sharing links: /:b:/s/sitename/...  (binary file sharing link)
      // These often don't reveal the extension, so we check the originating element
      // text for ".pdf" — handled separately in the click handler.

      return false;
    } catch {
      return false;
    }
  }

  /**
   * Builds a download URL from a SharePoint file URL.
   * Appends ?download=1 to force file download instead of browser preview.
   */
  function toDownloadUrl(url) {
    try {
      const u = new URL(url, window.location.origin);

      // If the URL has a .pdf path, just add download=1
      const pathname = decodeURIComponent(u.pathname).toLowerCase();
      if (pathname.endsWith('.pdf')) {
        u.searchParams.set('download', '1');
        return u.toString();
      }

      // If PDF path is in the 'id' parameter, construct a direct URL
      const idParam = u.searchParams.get('id');
      if (idParam && decodeURIComponent(idParam).toLowerCase().endsWith('.pdf')) {
        const directUrl = new URL(u.origin + idParam);
        directUrl.searchParams.set('download', '1');
        return directUrl.toString();
      }

      // Fallback: return as-is with download=1
      u.searchParams.set('download', '1');
      return u.toString();
    } catch {
      return url;
    }
  }

  /**
   * Extracts a human-readable filename from a URL.
   */
  function extractFilename(url) {
    try {
      const u = new URL(url, window.location.origin);

      // Try the path first
      const pathParts = u.pathname.split('/');
      for (let i = pathParts.length - 1; i >= 0; i--) {
        const part = decodeURIComponent(pathParts[i]);
        if (part.toLowerCase().endsWith('.pdf')) return part;
      }

      // Try the 'id' parameter
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

  /**
   * Checks if an element (or its visible text / aria-label) suggests a PDF.
   * Used for opaque sharing links where the URL doesn't reveal the file type.
   */
  function elementSuggestsPdf(el) {
    if (!el) return false;

    // Check text content of the element and immediate children
    const text = (el.textContent || '').trim().toLowerCase();
    if (text.endsWith('.pdf')) return true;

    // Check aria-label / title attributes
    const label = (el.getAttribute('aria-label') || '').toLowerCase();
    if (label.endsWith('.pdf')) return true;

    const title = (el.getAttribute('title') || '').toLowerCase();
    if (title.endsWith('.pdf')) return true;

    // Check data attributes used by SharePoint
    const dataFileName = el.getAttribute('data-filename') || el.getAttribute('data-file-name') || '';
    if (dataFileName.toLowerCase().endsWith('.pdf')) return true;

    return false;
  }

  // ── Click Interception ─────────────────────────────────────────────────

  /**
   * Walk up the DOM from the event target to find the nearest <a> link
   * or SharePoint file item element, up to `maxDepth` levels.
   */
  function findPdfLink(target, maxDepth = 15) {
    let el = target;
    for (let i = 0; i < maxDepth && el && el !== document.documentElement; i++) {
      // Standard <a> tag with href
      if (el.tagName === 'A' && el.href) {
        if (isPdfUrl(el.href)) {
          return { url: el.href, filename: extractFilename(el.href) };
        }
        // Check if the link text suggests PDF even if URL is opaque
        if (el.hostname?.endsWith('.sharepoint.com') && elementSuggestsPdf(el)) {
          return { url: el.href, filename: (el.textContent || '').trim() || extractFilename(el.href) };
        }
      }

      // SharePoint uses buttons for file names in modern list views
      if (el.tagName === 'BUTTON' || el.getAttribute('role') === 'link' || el.getAttribute('role') === 'gridcell') {
        if (elementSuggestsPdf(el)) {
          // Look for a sibling or parent element that has the actual URL
          const container = el.closest('[data-automationid="FieldRenderer-name"]')
            || el.closest('[data-automationid="name-column"]')
            || el.closest('.ms-DetailsRow');
          if (container) {
            const nearestLink = container.querySelector('a[href]');
            if (nearestLink && nearestLink.href) {
              return { url: nearestLink.href, filename: extractFilename(nearestLink.href) };
            }
          }
          // If no link found, the filename text itself is useful for URL-based detection
          // (the actual navigation will be caught by checkForViewerPage)
        }
      }

      el = el.parentElement;
    }
    return null;
  }

  /**
   * Main click handler — installed on document in CAPTURE phase so we
   * intercept before SharePoint's own handlers fire.
   */
  document.addEventListener('click', function (e) {
    if (!extensionEnabled) return;

    // Don't intercept if user is holding modifier keys (might want to open in new tab etc.)
    if (e.ctrlKey || e.shiftKey || e.metaKey || e.altKey) return;

    const pdfInfo = findPdfLink(e.target);
    if (!pdfInfo) return;

    // Prevent SharePoint from navigating to the PDF viewer
    e.preventDefault();
    e.stopImmediatePropagation();

    const downloadUrl = toDownloadUrl(pdfInfo.url);

    // Send to service worker for download + open in Acrobat
    chrome.runtime.sendMessage({
      type: 'OPEN_PDF',
      url: downloadUrl,
      filename: pdfInfo.filename
    }, (response) => {
      if (chrome.runtime.lastError) {
        console.warn('[SP PDF Opener] Failed to send message:', chrome.runtime.lastError.message);
        // Fallback: let the browser handle it normally
        window.location.href = pdfInfo.url;
      }
    });

    // Visual feedback: brief flash on the clicked element
    showInterceptFeedback(e.target);

  }, true); // ← capture phase

  // Also intercept auxclick (middle-click) for consistency
  document.addEventListener('auxclick', function (e) {
    if (!extensionEnabled) return;
    if (e.button !== 1) return; // middle click only

    const pdfInfo = findPdfLink(e.target);
    if (!pdfInfo) return;

    e.preventDefault();
    e.stopImmediatePropagation();

    const downloadUrl = toDownloadUrl(pdfInfo.url);
    chrome.runtime.sendMessage({
      type: 'OPEN_PDF',
      url: downloadUrl,
      filename: pdfInfo.filename
    });

    showInterceptFeedback(e.target);
  }, true);

  // ── Viewer Page Detection ──────────────────────────────────────────────

  /**
   * Track the last URL we processed to avoid duplicate interceptions.
   */
  let lastProcessedUrl = '';

  /**
   * Navigate to the parent directory in SharePoint's document library view.
   * Used when a PDF was opened in a new tab (from Excel, email, etc.)
   * and there's no browser history to go back to.
   *
   * @param {URL} u - The current page URL object
   * @param {string} decodedId - The decoded file path from the 'id' param
   */
  function navigateToParentDirectory(u, decodedId) {
    // Strategy 1: Use the 'parent' query param (SharePoint provides this)
    const parentParam = u.searchParams.get('parent');
    if (parentParam) {
      const parentUrl = new URL(u.origin + u.pathname);
      parentUrl.searchParams.set('id', decodeURIComponent(parentParam));
      parentUrl.searchParams.set('p', 'true');
      parentUrl.searchParams.set('ga', '1');
      console.log('[SP PDF Opener] Requesting parent directory navigation (from param)');
      chrome.runtime.sendMessage({ type: 'NAVIGATE_OR_FOCUS_PARENT', url: parentUrl.toString() });
      return;
    }

    // Strategy 2: Derive parent by stripping the filename from the id path
    if (decodedId) {
      const lastSlash = decodedId.lastIndexOf('/');
      if (lastSlash > 0) {
        const parentPath = decodedId.substring(0, lastSlash);
        const parentUrl = new URL(u.origin + u.pathname);
        parentUrl.searchParams.set('id', parentPath);
        parentUrl.searchParams.set('p', 'true');
        parentUrl.searchParams.set('ga', '1');
        console.log('[SP PDF Opener] Requesting parent directory navigation (derived)');
        chrome.runtime.sendMessage({ type: 'NAVIGATE_OR_FOCUS_PARENT', url: parentUrl.toString() });
        return;
      }
    }

    // Strategy 3: Can't determine parent — close the tab
    console.log('[SP PDF Opener] No parent directory found — closing tab');
    chrome.runtime.sendMessage({ type: 'CLOSE_TAB' });
  }

  /**
   * Detect if the current page/URL indicates a PDF is being viewed.
   * Handles multiple SharePoint URL patterns:
   *   1. AllItems.aspx?id=/path/to/file.pdf  (most common — document library inline viewer)
   *   2. WopiFrame.aspx / Doc.aspx           (Office online viewer)
   *   3. /:b:/ sharing links                 (sharing/preview links)
   */
  function checkForViewerPage() {
    if (!extensionEnabled) return;

    const url = window.location.href;

    // Debounce: don't process the same URL twice
    if (url === lastProcessedUrl) return;

    try {
      const u = new URL(url);

      // ── Pattern 1: AllItems.aspx?id=/path/to/file.pdf ──────────────
      // This is the MOST COMMON pattern in SharePoint Online document libraries.
      // When a user clicks a PDF in the file list, SharePoint navigates to:
      //   AllItems.aspx?...&id=%2Fsites%2F...%2Ffile%2Epdf&parentview=7
      const idParam = u.searchParams.get('id');
      if (idParam) {
        const decodedId = decodeURIComponent(idParam);
        if (decodedId.toLowerCase().endsWith('.pdf')) {
          console.log('[SP PDF Opener] Detected PDF in URL id param:', decodedId);
          lastProcessedUrl = url;

          // Construct direct download URL from the file path
          const downloadUrl = u.origin + decodedId + '?download=1';
          const filename = decodedId.split('/').pop() || 'document.pdf';

          chrome.runtime.sendMessage({
            type: 'OPEN_PDF',
            url: downloadUrl,
            filename: filename
          }, (response) => {
            if (chrome.runtime.lastError) {
              console.warn('[SP PDF Opener] Message failed:', chrome.runtime.lastError.message);
            } else {
              console.log('[SP PDF Opener] Download triggered for:', filename);
            }
          });

          // Navigate back to the document library (remove the id/parentview params)
          // Small delay to ensure the message is sent first
          setTimeout(() => {
            if (window.history.length > 1) {
              window.history.back();
            } else {
              // No history (opened from Excel, email, etc.) — navigate to parent directory
              navigateToParentDirectory(u, decodedId);
            }
          }, 300);

          return;
        }
      }

      // ── Pattern 2: WopiFrame / Doc.aspx viewer pages ───────────────
      if (url.includes('_layouts/15/WopiFrame.aspx') || url.includes('_layouts/15/Doc.aspx')) {
        const sourcedoc = u.searchParams.get('sourcedoc') || u.searchParams.get('SourceDoc');
        if (sourcedoc) {
          lastProcessedUrl = url;
          waitForElement('[data-automationid="download"]', (downloadBtn) => {
            const downloadHref = downloadBtn.getAttribute('href') || downloadBtn.closest('a')?.href;
            if (downloadHref) {
              chrome.runtime.sendMessage({
                type: 'OPEN_PDF',
                url: downloadHref,
                filename: document.title.replace(' - SharePoint', '').trim() || 'document.pdf'
              });
              if (window.history.length > 1) {
                window.history.back();
              } else {
                chrome.runtime.sendMessage({ type: 'CLOSE_TAB' });
              }
            }
          });
        }
      }

      // ── Pattern 3: Sharing link viewer (/:b:/) ─────────────────────
      if (/\/:b:\//.test(url)) {
        lastProcessedUrl = url;
        waitForElement('[data-automationid="download"]', (downloadBtn) => {
          const downloadHref = downloadBtn.getAttribute('href') || downloadBtn.closest('a')?.href;
          if (downloadHref && isPdfUrl(downloadHref)) {
            chrome.runtime.sendMessage({
              type: 'OPEN_PDF',
              url: toDownloadUrl(downloadHref),
              filename: document.title.replace(' - SharePoint', '').trim() || 'document.pdf'
            });
            if (window.history.length > 1) {
              window.history.back();
            } else {
              chrome.runtime.sendMessage({ type: 'CLOSE_TAB' });
            }
          }
        });
      }

    } catch (err) {
      console.warn('[SP PDF Opener] Viewer detection error:', err);
    }
  }

  /**
   * Wait for a DOM element to appear (SharePoint loads asynchronously).
   */
  function waitForElement(selector, callback, timeout = 5000) {
    const existing = document.querySelector(selector);
    if (existing) {
      callback(existing);
      return;
    }

    const observer = new MutationObserver(() => {
      const el = document.querySelector(selector);
      if (el) {
        observer.disconnect();
        callback(el);
      }
    });

    observer.observe(document.documentElement, { childList: true, subtree: true });

    // Timeout safety
    setTimeout(() => observer.disconnect(), timeout);
  }

  // Run viewer detection when the page is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', checkForViewerPage);
  } else {
    checkForViewerPage();
  }

  // ── Visual Feedback ────────────────────────────────────────────────────

  /**
   * Shows a brief visual pulse on the intercepted element so the user
   * knows their click was registered.
   */
  function showInterceptFeedback(target) {
    // Find the nearest visible text element
    let el = target;
    for (let i = 0; i < 5 && el; i++) {
      if (el.offsetWidth > 0 && el.offsetHeight > 0) break;
      el = el.parentElement;
    }
    if (!el) return;

    const originalOutline = el.style.outline;
    const originalTransition = el.style.transition;

    el.style.transition = 'outline-color 0.3s ease';
    el.style.outline = '2px solid rgba(233, 69, 96, 0.8)';

    setTimeout(() => {
      el.style.outline = '2px solid rgba(233, 69, 96, 0)';
      setTimeout(() => {
        el.style.outline = originalOutline;
        el.style.transition = originalTransition;
      }, 300);
    }, 400);
  }

  // ── SPA Navigation Monitoring ──────────────────────────────────────────

  /**
   * SharePoint is a SPA — pages change without full reloads.
   * Monitor URL changes to re-run viewer detection.
   */
  let lastUrl = window.location.href;

  const urlObserver = new MutationObserver(() => {
    if (window.location.href !== lastUrl) {
      lastUrl = window.location.href;
      checkForViewerPage();
    }
  });

  urlObserver.observe(document.documentElement, { childList: true, subtree: true });

  // Also listen for pushState/replaceState (which SharePoint uses)
  const originalPushState = history.pushState;
  const originalReplaceState = history.replaceState;

  history.pushState = function (...args) {
    originalPushState.apply(this, args);
    setTimeout(checkForViewerPage, 100);
  };

  history.replaceState = function (...args) {
    originalReplaceState.apply(this, args);
    setTimeout(checkForViewerPage, 100);
  };

  window.addEventListener('popstate', () => {
    setTimeout(checkForViewerPage, 100);
  });

})();
