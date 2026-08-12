/**
 * SP PDF Opener - Popup Logic
 * Handles popup state, UI updates, and chrome messaging/storage synchronization.
 */

document.addEventListener('DOMContentLoaded', () => {
  const toggleInput = document.getElementById('toggle-enabled');
  const statusDot = document.getElementById('status-dot');
  const statusText = document.getElementById('status-text');
  const pdfCountElem = document.getElementById('pdf-count');
  const openDownloadsBtn = document.getElementById('open-downloads');

  /**
   * Updates the status banner UI based on enabled flag.
   * @param {boolean} isEnabled
   */
  const updateStatusUI = (isEnabled) => {
    if (isEnabled) {
      statusText.textContent = 'Active';
      statusText.classList.remove('paused');
      statusText.classList.add('active');

      statusDot.classList.remove('paused');
      statusDot.classList.add('active');
    } else {
      statusText.textContent = 'Paused';
      statusText.classList.remove('active');
      statusText.classList.add('paused');

      statusDot.classList.remove('active');
      statusDot.classList.add('paused');
    }
  };

  /**
   * Initializes state from chrome.storage.local.
   */
  chrome.storage.local.get({ enabled: true, pdfCount: 0 }, (items) => {
    toggleInput.checked = items.enabled;
    updateStatusUI(items.enabled);
    pdfCountElem.textContent = items.pdfCount;
  });

  /**
   * Handle toggle switch changes.
   */
  toggleInput.addEventListener('change', () => {
    const isEnabled = toggleInput.checked;
    updateStatusUI(isEnabled);

    // Save state to chrome.storage.local
    chrome.storage.local.set({ enabled: isEnabled });

    // Send message to background script
    chrome.runtime.sendMessage({
      type: 'TOGGLE_ENABLED',
      enabled: isEnabled
    });
  });

  /**
   * Handle "Open downloads folder" link click.
   */
  openDownloadsBtn.addEventListener('click', (event) => {
    event.preventDefault();
    chrome.runtime.sendMessage({
      type: 'OPEN_DOWNLOADS'
    });
  });

  /**
   * Listen for real-time changes in chrome.storage.local
   */
  chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName !== 'local') return;

    if (changes.pdfCount) {
      pdfCountElem.textContent = changes.pdfCount.newValue ?? 0;
    }

    if (changes.enabled) {
      toggleInput.checked = changes.enabled.newValue;
      updateStatusUI(changes.enabled.newValue);
    }
  });
});
