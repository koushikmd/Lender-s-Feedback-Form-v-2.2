"""
Lender Feedback Tool - Desktop Edition.
Runs a local Flask server on 127.0.0.1 (localhost only) and opens in the default browser.
Data never leaves the user's machine.
"""
import os
import sys
import io
import json
import tempfile
import webbrowser
import threading
import logging
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template_string


# Suppress Flask's default request logging
log = logging.getLogger('werkzeug')
log.setLevel(logging.WARNING)


def resource_path(rel_path):
    """Resolve path for PyInstaller-bundled resources."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, rel_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), rel_path)


# Add bundled module paths so imports work both from source and from PyInstaller bundle
if hasattr(sys, '_MEIPASS'):
    sys.path.insert(0, sys._MEIPASS)
else:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pdf_extractor import extract_all_data
from docx_generator import generate_docx


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB


HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Lender Feedback Form Tool</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root {
  --bg: #0f1419;
  --panel: #1a1f2e;
  --panel-light: #232937;
  --border: #2d3548;
  --text: #e4e7eb;
  --text-muted: #9ca3af;
  --accent: #4A7BF7;
  --accent-hover: #5a8bff;
  --success: #34C77B;
  --warning: #F5A623;
  --error: #ef4444;
  --appraisal: #4A7BF7;
  --isbs: #34C77B;
  --manual: #F5A623;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: 'DM Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
  padding: 24px;
  line-height: 1.5;
}
.container { max-width: 1200px; margin: 0 auto; }
header {
  text-align: center;
  padding: 32px 0 24px;
  border-bottom: 1px solid var(--border);
  margin-bottom: 32px;
}
h1 {
  font-size: 28px;
  font-weight: 700;
  letter-spacing: -0.5px;
  margin-bottom: 8px;
}
.subtitle { color: var(--text-muted); font-size: 14px; }
.offline-badge {
  display: inline-block;
  background: rgba(52, 199, 123, 0.15);
  color: var(--success);
  padding: 4px 12px;
  border-radius: 12px;
  font-size: 12px;
  font-weight: 500;
  margin-top: 8px;
  border: 1px solid rgba(52, 199, 123, 0.3);
}
.steps { display: flex; gap: 8px; margin-bottom: 24px; }
.step {
  flex: 1;
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 12px 16px;
  font-size: 14px;
  color: var(--text-muted);
}
.step.active { border-color: var(--accent); color: var(--text); background: var(--panel-light); }
.step.done { border-color: var(--success); color: var(--success); }
.panel {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 24px;
  margin-bottom: 16px;
}
.panel h2 { font-size: 18px; font-weight: 600; margin-bottom: 16px; }
.upload-zone { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 24px; }
.dropzone {
  border: 2px dashed var(--border);
  border-radius: 8px;
  padding: 32px 16px;
  text-align: center;
  cursor: pointer;
  transition: all 0.2s;
  background: var(--panel-light);
}
.dropzone:hover { border-color: var(--accent); background: rgba(74, 123, 247, 0.05); }
.dropzone.has-file { border-color: var(--success); background: rgba(52, 199, 123, 0.05); }
.dropzone.dragover { border-color: var(--accent); background: rgba(74, 123, 247, 0.1); }
.dropzone-label { font-size: 14px; font-weight: 500; margin-bottom: 8px; }
.dropzone-hint { font-size: 12px; color: var(--text-muted); }
.dropzone input[type=file] { display: none; }
.dropzone .filename {
  font-family: 'JetBrains Mono', monospace;
  font-size: 12px;
  color: var(--success);
  margin-top: 8px;
  word-break: break-all;
}
.btn {
  padding: 10px 20px;
  border-radius: 6px;
  border: none;
  cursor: pointer;
  font-weight: 500;
  font-size: 14px;
  font-family: inherit;
  transition: all 0.2s;
}
.btn-primary { background: var(--accent); color: white; }
.btn-primary:hover:not(:disabled) { background: var(--accent-hover); }
.btn-primary:disabled { background: var(--border); color: var(--text-muted); cursor: not-allowed; }
.btn-secondary { background: transparent; color: var(--text); border: 1px solid var(--border); }
.btn-secondary:hover { border-color: var(--accent); color: var(--accent); }
.btn-success { background: var(--success); color: white; }
.btn-success:hover { background: #2da86a; }
.actions { display: flex; gap: 12px; justify-content: flex-end; margin-top: 16px; }
.review-section { margin-bottom: 24px; }
.review-section-header {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 12px;
  padding-bottom: 8px;
  border-bottom: 1px solid var(--border);
}
.review-section-header h3 { font-size: 15px; font-weight: 600; }
.source-badge {
  font-size: 11px;
  padding: 2px 8px;
  border-radius: 10px;
  font-weight: 500;
}
.source-appraisal { background: rgba(74, 123, 247, 0.15); color: var(--appraisal); }
.source-isbs { background: rgba(52, 199, 123, 0.15); color: var(--isbs); }
.source-manual { background: rgba(245, 166, 35, 0.15); color: var(--manual); }
.field {
  display: grid;
  grid-template-columns: 180px 1fr;
  gap: 12px;
  margin-bottom: 10px;
  align-items: start;
}
.field label {
  font-size: 13px;
  color: var(--text-muted);
  padding-top: 8px;
}
.field input, .field textarea, .field select {
  background: var(--panel-light);
  border: 1px solid var(--border);
  color: var(--text);
  padding: 8px 10px;
  border-radius: 6px;
  font-size: 13px;
  font-family: inherit;
  width: 100%;
}
.field input:focus, .field textarea:focus { outline: none; border-color: var(--accent); }
.field textarea { resize: vertical; min-height: 70px; font-family: inherit; }
.liability-table { overflow-x: auto; margin-bottom: 12px; }
.liability-table table { width: 100%; border-collapse: collapse; font-size: 12px; }
.liability-table th, .liability-table td {
  padding: 6px 8px;
  text-align: left;
  border: 1px solid var(--border);
}
.liability-table th {
  background: var(--panel-light);
  font-weight: 600;
  color: var(--text-muted);
  font-size: 11px;
}
.liability-table input {
  background: transparent;
  border: none;
  color: var(--text);
  width: 100%;
  font-size: 12px;
  font-family: inherit;
  padding: 2px 4px;
}
.liability-table input:focus { background: var(--panel-light); outline: 1px solid var(--accent); }
.liability-table .remove-btn {
  background: transparent;
  border: none;
  color: var(--error);
  cursor: pointer;
  padding: 4px 8px;
  font-size: 16px;
}
.hidden { display: none !important; }
.error-msg {
  background: rgba(239, 68, 68, 0.1);
  color: var(--error);
  border: 1px solid rgba(239, 68, 68, 0.3);
  padding: 12px 16px;
  border-radius: 6px;
  margin: 16px 0;
  font-size: 13px;
}
.spinner {
  display: inline-block;
  width: 16px; height: 16px;
  border: 2px solid var(--border);
  border-top-color: var(--accent);
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
  vertical-align: middle;
}
@keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>
<div class="container">
  <header>
    <h1>Lender's Feedback Form Tool</h1>
    <div class="subtitle">IDLC Finance Limited &mdash; Credit Rating Document Automation</div>
    <div class="offline-badge">🔒 Fully Offline &mdash; Running on localhost, data never leaves your computer</div>
  </header>

  <div class="steps">
    <div class="step active" id="step1">1. Upload PDFs</div>
    <div class="step" id="step2">2. Review &amp; Edit</div>
    <div class="step" id="step3">3. Download DOCX</div>
  </div>

  <div id="uploadPanel" class="panel">
    <h2>Upload Source Documents</h2>
    <div class="upload-zone">
      <label class="dropzone" id="appraisalZone">
        <div class="dropzone-label">Appraisal PDF</div>
        <div class="dropzone-hint">Click or drag PDF here</div>
        <div class="filename" id="appraisalName"></div>
        <input type="file" id="appraisalFile" accept=".pdf">
      </label>
      <label class="dropzone" id="isbsZone">
        <div class="dropzone-label">ISBS PDF</div>
        <div class="dropzone-hint">Click or drag PDF here</div>
        <div class="filename" id="isbsName"></div>
        <input type="file" id="isbsFile" accept=".pdf">
      </label>
    </div>
    <div class="actions">
      <button class="btn btn-primary" id="extractBtn" disabled>Extract Data</button>
    </div>
    <div id="uploadError" class="error-msg hidden"></div>
  </div>

  <div id="reviewPanel" class="panel hidden">
    <h2>Review &amp; Edit Extracted Data</h2>
    <div id="reviewContent"></div>
    <div class="actions">
      <button class="btn btn-secondary" id="backBtn">Back</button>
      <button class="btn btn-success" id="generateBtn">Generate DOCX</button>
    </div>
    <div id="reviewError" class="error-msg hidden"></div>
  </div>

  <div id="downloadPanel" class="panel hidden">
    <h2>Document Ready</h2>
    <p style="color: var(--text-muted); margin-bottom: 16px;">Your Lender's Feedback Form has been generated. Click below to download.</p>
    <div class="actions">
      <button class="btn btn-secondary" id="startOverBtn">Start Over</button>
      <button class="btn btn-success" id="downloadBtn">Download DOCX</button>
    </div>
  </div>
</div>

<script>
let appraisalFile = null;
let isbsFile = null;
let extractedData = null;

const $ = id => document.getElementById(id);

function setStep(n) {
  for (let i = 1; i <= 3; i++) {
    const el = $('step' + i);
    el.classList.remove('active', 'done');
    if (i < n) el.classList.add('done');
    if (i === n) el.classList.add('active');
  }
}
function showPanel(which) {
  ['uploadPanel','reviewPanel','downloadPanel'].forEach(p => {
    $(p).classList.toggle('hidden', p !== which);
  });
}
function showError(panelId, msg) {
  const el = $(panelId);
  el.textContent = msg;
  el.classList.remove('hidden');
}
function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function setupDropzone(zoneId, inputId, nameId, setter) {
  const zone = $(zoneId);
  const input = $(inputId);
  const nameEl = $(nameId);
  const handle = (file) => {
    if (!file || !file.name.toLowerCase().endsWith('.pdf')) return;
    setter(file);
    nameEl.textContent = file.name;
    zone.classList.add('has-file');
    checkReady();
  };
  input.addEventListener('change', e => handle(e.target.files[0]));
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('dragover');
    handle(e.dataTransfer.files[0]);
  });
}

setupDropzone('appraisalZone', 'appraisalFile', 'appraisalName', f => appraisalFile = f);
setupDropzone('isbsZone', 'isbsFile', 'isbsName', f => isbsFile = f);

function checkReady() {
  $('extractBtn').disabled = !(appraisalFile && isbsFile);
}

$('extractBtn').addEventListener('click', async () => {
  $('extractBtn').disabled = true;
  $('extractBtn').innerHTML = '<span class="spinner"></span> Extracting...';
  $('uploadError').classList.add('hidden');

  try {
    const fd = new FormData();
    fd.append('appraisal', appraisalFile);
    fd.append('isbs', isbsFile);
    const res = await fetch('/api/extract', { method: 'POST', body: fd });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: 'Extraction failed' }));
      throw new Error(err.error || 'Extraction failed');
    }
    extractedData = await res.json();
    renderReview();
    showPanel('reviewPanel');
    setStep(2);
  } catch (err) {
    showError('uploadError', err.message);
  } finally {
    $('extractBtn').disabled = false;
    $('extractBtn').textContent = 'Extract Data';
  }
});

function renderReview() {
  const d = extractedData;
  const html = `
    <div class="review-section">
      <div class="review-section-header">
        <h3>General Information</h3>
        <span class="source-badge source-appraisal">From Appraisal</span>
      </div>
      ${fieldHtml('Business Name', 'business_name', d.business_name)}
      ${fieldHtml('Address', 'address', d.address)}
      ${fieldHtml('Key Person', 'key_person', d.key_person)}
      ${fieldHtml('Contact Number', 'contact_number', d.contact_number)}
      ${fieldHtml('Ownership Type', 'ownership_type', d.ownership_type)}
      ${fieldHtml('RM Name', 'rm_name', d.rm_name)}
      ${fieldHtml('RM Contact', 'rm_contact', d.rm_contact || '', 'manual')}
      ${fieldHtml('Date', 'date', d.date)}
    </div>
    <div class="review-section">
      <div class="review-section-header">
        <h3>Business Dynamics</h3>
        <span class="source-badge source-appraisal">From Appraisal</span>
      </div>
      ${textareaHtml('Background', 'background', d.background)}
      ${textareaHtml('Major Suppliers', 'major_suppliers', d.major_suppliers)}
      ${textareaHtml('Major Clients', 'major_clients', d.major_clients)}
    </div>
    <div class="review-section">
      <div class="review-section-header">
        <h3>Financial Information</h3>
        <span class="source-badge source-isbs">From ISBS</span>
      </div>
      ${fieldHtml('Sales / Revenue', 'sales_revenue', d.sales_revenue)}
      ${fieldHtml('Inventory', 'inventory', d.inventory)}
      ${fieldHtml('A/C Receivable', 'ar', d.ar)}
      ${fieldHtml('A/C Payable', 'ap', d.ap)}
      ${fieldHtml('Gross Profit (%)', 'gp_pct', d.gp_pct)}
      ${fieldHtml('Net Profit (%)', 'np_pct', d.np_pct)}
    </div>
    <div class="review-section">
      <div class="review-section-header">
        <h3>Long Term Liabilities</h3>
        <span class="source-badge source-appraisal">From Appraisal</span>
      </div>
      <div class="liability-table" id="ltTable"></div>
      <button class="btn btn-secondary" onclick="addRow('lt')" style="font-size:12px;padding:6px 12px;">+ Add LT Row</button>
    </div>
    <div class="review-section">
      <div class="review-section-header">
        <h3>Short Term Liabilities</h3>
        <span class="source-badge source-appraisal">From Appraisal</span>
      </div>
      <div class="liability-table" id="stTable"></div>
      <button class="btn btn-secondary" onclick="addRow('st')" style="font-size:12px;padding:6px 12px;">+ Add ST Row</button>
    </div>
    <div class="review-section">
      <div class="review-section-header">
        <h3>Import/Export &amp; Credit</h3>
      </div>
      ${textareaHtml('Import/Export', 'import_export', d.import_export)}
      ${fieldHtml('Total Credit Summation', 'credit_summation', d.credit_summation)}
    </div>
  `;
  $('reviewContent').innerHTML = html;
  renderLiabilityTable('lt');
  renderLiabilityTable('st');
}

function fieldHtml(label, key, value, src) {
  const badge = src === 'manual' ? '<span class="source-badge source-manual" style="margin-left:8px;">Manual</span>' : '';
  return `
    <div class="field">
      <label>${label} ${badge}</label>
      <input type="text" data-field="${key}" value="${escapeHtml(value || '')}" />
    </div>
  `;
}
function textareaHtml(label, key, value) {
  return `
    <div class="field">
      <label>${label}</label>
      <textarea data-field="${key}">${escapeHtml(value || '')}</textarea>
    </div>
  `;
}

function renderLiabilityTable(type) {
  const isLt = type === 'lt';
  const containerId = isLt ? 'ltTable' : 'stTable';
  const rows = isLt ? extractedData.long_term_liabilities : extractedData.short_term_liabilities;
  const cols = isLt
    ? [['bank_name','Bank Name'],['facility_type','Facility Type'],['limit','Limit'],['outstanding','Outstanding'],['emi','EMI'],['tenure','Tenure'],['repayment_status','Status']]
    : [['bank_name','Bank Name'],['facility_type','Facility Type'],['limit','Limit'],['outstanding','Outstanding'],['recycle_times','Recycle'],['expiry_date','Expiry'],['repayment_status','Status']];

  let html = '<table><thead><tr>';
  cols.forEach(c => { html += `<th>${c[1]}</th>`; });
  html += '<th></th></tr></thead><tbody>';
  rows.forEach((row, idx) => {
    html += '<tr>';
    cols.forEach(c => {
      html += `<td><input type="text" data-list="${type}" data-idx="${idx}" data-key="${c[0]}" value="${escapeHtml(row[c[0]] || '')}" /></td>`;
    });
    html += `<td><button class="remove-btn" onclick="removeRow('${type}', ${idx})">×</button></td>`;
    html += '</tr>';
  });
  html += '</tbody></table>';
  $(containerId).innerHTML = html;
}

window.addRow = function(type) {
  const emptyLt = { bank_name:'', facility_type:'', limit:'', outstanding:'', emi:'', tenure:'', repayment_status:'Regular' };
  const emptySt = { bank_name:'', facility_type:'', limit:'', outstanding:'', recycle_times:'', expiry_date:'', repayment_status:'Regular' };
  collectFormData();
  if (type === 'lt') extractedData.long_term_liabilities.push(emptyLt);
  else extractedData.short_term_liabilities.push(emptySt);
  renderLiabilityTable(type);
};
window.removeRow = function(type, idx) {
  collectFormData();
  if (type === 'lt') extractedData.long_term_liabilities.splice(idx, 1);
  else extractedData.short_term_liabilities.splice(idx, 1);
  renderLiabilityTable(type);
};

function collectFormData() {
  document.querySelectorAll('[data-field]').forEach(el => {
    extractedData[el.getAttribute('data-field')] = el.value;
  });
  document.querySelectorAll('[data-list]').forEach(el => {
    const list = el.getAttribute('data-list') === 'lt'
      ? extractedData.long_term_liabilities
      : extractedData.short_term_liabilities;
    const idx = parseInt(el.getAttribute('data-idx'));
    const key = el.getAttribute('data-key');
    if (list[idx]) list[idx][key] = el.value;
  });
}

$('generateBtn').addEventListener('click', async () => {
  $('generateBtn').disabled = true;
  $('generateBtn').innerHTML = '<span class="spinner"></span> Generating...';
  $('reviewError').classList.add('hidden');

  try {
    collectFormData();
    const res = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(extractedData)
    });
    if (!res.ok) throw new Error('Generation failed');
    const blob = await res.blob();
    window._docxBlob = blob;
    showPanel('downloadPanel');
    setStep(3);
  } catch (err) {
    showError('reviewError', err.message);
  } finally {
    $('generateBtn').disabled = false;
    $('generateBtn').textContent = 'Generate DOCX';
  }
});

$('downloadBtn').addEventListener('click', () => {
  if (!window._docxBlob) return;
  const url = URL.createObjectURL(window._docxBlob);
  const a = document.createElement('a');
  const businessName = (extractedData.business_name || 'Client').replace(/[^\w\s-]/g, '').trim().replace(/\s+/g, '_');
  a.href = url;
  a.download = `Lender_Feedback_${businessName}.docx`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
});

$('backBtn').addEventListener('click', () => {
  showPanel('uploadPanel');
  setStep(1);
});

$('startOverBtn').addEventListener('click', () => {
  appraisalFile = null;
  isbsFile = null;
  extractedData = null;
  window._docxBlob = null;
  $('appraisalName').textContent = '';
  $('isbsName').textContent = '';
  $('appraisalZone').classList.remove('has-file');
  $('isbsZone').classList.remove('has-file');
  $('appraisalFile').value = '';
  $('isbsFile').value = '';
  checkReady();
  showPanel('uploadPanel');
  setStep(1);
});
</script>
</body>
</html>
"""


@app.route('/')
def index():
    return render_template_string(HTML_PAGE)


@app.route('/api/extract', methods=['POST'])
def extract():
    try:
        appraisal_file = request.files.get('appraisal')
        isbs_file = request.files.get('isbs')
        if not appraisal_file or not isbs_file:
            return jsonify({'error': 'Both PDF files are required'}), 400

        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as ap_tmp:
            appraisal_file.save(ap_tmp.name)
            ap_path = ap_tmp.name
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as is_tmp:
            isbs_file.save(is_tmp.name)
            is_path = is_tmp.name

        try:
            data = extract_all_data(ap_path, is_path)
        finally:
            try: os.unlink(ap_path)
            except OSError: pass
            try: os.unlink(is_path)
            except OSError: pass

        return jsonify(data)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json()
        docx_bytes = generate_docx(data)
        buf = io.BytesIO(docx_bytes)
        buf.seek(0)
        return send_file(
            buf,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='Lender_Feedback_Form.docx'
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/health')
def health():
    return jsonify({'status': 'ok'})


def find_free_port(start=5050, end=5100):
    import socket
    for port in range(start, end):
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.bind(('127.0.0.1', port))
            s.close()
            return port
        except OSError:
            s.close()
            continue
    return start


def open_browser(url, delay=1.2):
    import time
    time.sleep(delay)
    try:
        webbrowser.open(url)
    except Exception:
        pass


def main():
    port = find_free_port()
    url = f'http://127.0.0.1:{port}/'

    print("=" * 60)
    print("Lender Feedback Form Tool")
    print("IDLC Finance Limited - Desktop Edition")
    print("=" * 60)
    print(f"\nStarting local server at {url}")
    print("\n(Data never leaves your computer - all processing is local)")
    print("\nThe tool will open in your default browser in a moment...")
    print("\nTo stop the tool, close this window.\n")

    # Open browser in a separate thread after the server starts
    threading.Thread(target=open_browser, args=(url,), daemon=True).start()

    # Bind to 127.0.0.1 ONLY (localhost) - never accessible from network
    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)


if __name__ == '__main__':
    main()
