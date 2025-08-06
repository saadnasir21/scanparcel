const fs = require('fs');
const path = require('path');
const assert = require('assert');
const { JSDOM } = require('jsdom');

const html = fs.readFileSync(path.join(__dirname, '..', 'ScannerSidebar.html'), 'utf8');
const dom = new JSDOM(html, { runScripts: 'dangerously', resources: 'usable' });

// wait for DOMContentLoaded
function runTest() {
  const { window } = dom;
  const { document } = window;
  const calls = {};
  window.google = {
    script: {
      run: {
        withSuccessHandler(cb) {
          calls.successHandler = cb;
          return {
            processParcelScan(code) {
              calls.processCode = code;
              // do nothing to simulate pending server response
            }
          };
        }
      }
    }
  };

  const input = document.getElementById('parcelInput');
  const status = document.getElementById('statusMessage');
  input.value = 'ABC123';
  window.submitScan();

  try {
    assert.strictEqual(input.value, '', 'input should be cleared immediately');
    assert.strictEqual(status.textContent, 'Processing…', 'status should show Processing…');
    console.log('✓ immediate clear and status update');
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }
}

if (dom.window.document.readyState === 'complete') {
  runTest();
} else {
  dom.window.addEventListener('load', runTest);
}
