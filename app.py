"""
Render backend for the Examination Tool.

Reuses extractor.py's charge-sheet reading, classification, and document
generation logic exactly as built and tested in the desktop app — nothing
in that logic was changed. This file only adds: a web upload form in place
of the Tkinter window, a server-side Anthropic API key (so individual
courts never see or handle a key), and the same credits-based license
check already used for Result Portions, pointed at its own collection.

Multi-user safety: each request gets its own temp directory and its own
in-memory result list — no global mutable state, so two courts uploading
at the same moment never interfere with each other, regardless of whether
they land on the same or different server instances.
"""

import os
import io
import sys
import time
import zipfile
import tempfile
import shutil
from pathlib import Path

from flask import Flask, request, send_file, jsonify, render_template_string

from extractor import process_file
from credits import check_credits, deduct_credits, InsufficientCredits, LicenseNotFound, LicenseInactive

app = Flask(__name__)

MAX_FILES = 20  # a practical per-batch ceiling, comfortably inside the request timeout
SERVER_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", file=sys.stderr, flush=True)


# --- Firebase Admin init -----------------------------------------------
# Same pattern as the Result Portions backend: set FIREBASE_SERVICE_ACCOUNT_JSON
# as an environment variable on Render (the full JSON content of a Firebase
# service account key, from Firebase Console -> Project Settings -> Service
# Accounts -> Generate new private key).
_db = None


def get_db():
    global _db
    if _db is not None:
        return _db
    import firebase_admin
    from firebase_admin import credentials, firestore as fb_firestore

    if not firebase_admin._apps:
        cred_json = os.environ.get("FIREBASE_SERVICE_ACCOUNT_JSON")
        if cred_json:
            import json
            cred = credentials.Certificate(json.loads(cred_json))
            firebase_admin.initialize_app(cred)
        else:
            firebase_admin.initialize_app()
    _db = fb_firestore.client()
    return _db


@app.route("/health")
def health():
    return jsonify({"status": "ok", "api_key_configured": bool(SERVER_API_KEY)})


@app.route("/api/process", methods=["POST"])
def api_process():
    log("=== new request received ===")

    if not SERVER_API_KEY:
        return jsonify({"error": "Server is not configured with an API key yet. Contact support."}), 500

    court = request.form.get("court", "").strip()
    judge = request.form.get("judge", "").strip()
    court_type = request.form.get("court_type", "magistrate").strip()
    mode = request.form.get("mode", "auto").strip()
    manual_ref_code = request.form.get("manual_ref", "").strip()
    license_key = request.form.get("license_key", "").strip()
    uploaded = request.files.getlist("files")

    log(f"court={court!r}, court_type={court_type}, mode={mode}, files={len(uploaded)}")

    if not court:
        return jsonify({"error": "Court name is required."}), 400
    if not license_key:
        return jsonify({"error": "Missing license key."}), 400
    if not uploaded:
        return jsonify({"error": "No files were uploaded."}), 400
    if len(uploaded) > MAX_FILES:
        return jsonify({"error": f"Too many files in one batch (max {MAX_FILES}). Please split into smaller batches."}), 400

    # Resolve manual_ref the same way the desktop dialog did
    manual_ref = None
    if court_type == "magistrate" and mode == "manual" and manual_ref_code:
        manual_ref = manual_ref_code + " Cr.P.C." if manual_ref_code in ("251", "239") else manual_ref_code + " BNSS"

    cfg = {"court": court, "judge": judge, "api_key": SERVER_API_KEY}

    workdir = tempfile.mkdtemp(prefix="exam_")
    log(f"workdir={workdir}")
    try:
        saved_paths = []
        for f in uploaded:
            ext = os.path.splitext(f.filename)[1].lower()
            if ext not in (".pdf", ".docx", ".doc"):
                log(f"skipping unsupported file: {f.filename}")
                continue
            safe_path = os.path.join(workdir, os.path.basename(f.filename))
            f.save(safe_path)
            saved_paths.append(safe_path)

        if not saved_paths:
            return jsonify({"error": "No supported files (.pdf, .docx, .doc) were found in the upload."}), 400

        log(f"checking credits for {license_key}, {len(saved_paths)} file(s)...")
        try:
            check_credits(get_db(), license_key, len(saved_paths))
        except LicenseNotFound as e:
            return jsonify({"error": str(e)}), 403
        except LicenseInactive as e:
            return jsonify({"error": str(e)}), 403
        except InsufficientCredits as e:
            return jsonify({"error": str(e)}), 402
        log("credits check passed (not yet deducted — happens after the batch finishes)")

        results = []
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in saved_paths:
                fname = os.path.basename(path)
                log(f"processing {fname} ...")
                t0 = time.time()
                try:
                    result = process_file(path, cfg, court_type, manual_ref, progress_cb=None)
                    for d in result["docs"]:
                        zf.writestr(d["fname"], d["data"])
                    results.append({"name": fname, "status": "ok", "ct": result["ct"]})
                    log(f"done {fname} in {round(time.time()-t0,2)}s -> {result['ct']}")
                except Exception as e:
                    log(f"FAILED {fname}: {e}")
                    results.append({"name": fname, "status": "failed", "error": str(e)})

        ok_count = sum(1 for r in results if r["status"] == "ok")
        if ok_count == 0:
            return jsonify({"error": "None of the uploaded files could be processed.", "results": results}), 400

        # Only now, with processing genuinely finished, are credits deducted —
        # counted per file actually attempted (matching "Rs.20 per examination"),
        # same two-phase check-then-deduct split used for Result Portions.
        try:
            new_remaining = deduct_credits(get_db(), license_key, len(saved_paths))
            log(f"credits deducted, {new_remaining} remaining")
        except Exception as e:
            log(f"WARNING: credit deduction failed after successful processing: {e}")

        zip_buffer.seek(0)
        response = send_file(
            zip_buffer, mimetype="application/zip", as_attachment=True,
            download_name="Examination documents.zip",
        )
        response.headers["X-Results"] = "; ".join(f'{r["name"]}:{r["status"]}' for r in results)
        return response

    except Exception as e:
        return jsonify({"error": f"Unexpected error: {e}"}), 500
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


@app.route("/")
def index():
    return render_template_string(PAGE)


PAGE = """<!DOCTYPE html>
<html>
<head>
<script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-9117928033972031"
     crossorigin="anonymous"></script>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Examination Tool</title>
<style>
  :root { --navy:#1a237e; --navy-bg:#e8eaf6; --green:#1e7a3e; --green-bg:#e8f5e9;
          --red:#c0392b; --red-bg:#fdecea; --border:#dcdcdc; }
  body { font-family: Arial, sans-serif; max-width: 720px; margin: 0 auto; padding: 30px 24px 60px;
         color:#1c2530; background:#fafafa; }
  .card { background:white; border:1px solid var(--border); border-radius:10px; padding:24px 28px;
          margin-bottom:20px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }
  h1 { font-size:21px; margin:0 0 4px; }
  .subtitle { color:#666; font-size:13px; margin:0 0 20px; }
  .notice { background:var(--navy-bg); color:var(--navy); border-radius:8px; padding:12px 14px;
            font-size:13px; margin-bottom:18px; line-height:1.5; }
  label { display:block; margin-top:16px; margin-bottom:6px; font-weight:bold; font-size:14px; }
  input[type=text], input[type=file] { width:100%; padding:9px 10px; border:1px solid var(--border);
         border-radius:6px; font-size:14px; background:#fcfcfc; box-sizing:border-box; }
  .radio-row { display:flex; gap:18px; margin-top:6px; flex-wrap:wrap; }
  .radio-row label { display:flex; align-items:center; gap:6px; font-weight:normal; margin-top:0; font-size:13.5px; }
  #manualBlock, #modeBlock { display:none; }
  #runBtn { margin-top:22px; padding:12px 28px; background:var(--navy); color:white; border:none;
            border-radius:6px; font-size:15px; font-weight:bold; cursor:pointer; }
  #runBtn:disabled { background:#a5a5a5; cursor:default; }
  #resultBox { margin-top:18px; padding:14px 16px; border-radius:8px; font-size:13.5px; display:none; white-space:pre-line; }
  .ok { background:var(--green-bg); color:var(--green); }
  .err { background:var(--red-bg); color:var(--red); }
</style>
</head>
<body>

<!-- ── LICENSE SCREEN ── -->
<style>
#examLicScreen{position:fixed;top:0;left:0;width:100%;height:100%;background:#1a237e;display:flex;align-items:center;justify-content:center;z-index:99999;font-family:Arial,sans-serif}
.exam-lic-box{background:#fff;border-radius:14px;padding:40px;max-width:440px;width:90%;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.35)}
.exam-lic-icon{font-size:40px;margin-bottom:14px}
.exam-lic-title{font-size:19px;font-weight:700;color:#1a237e;margin-bottom:6px}
.exam-lic-sub{font-size:13px;color:#666;margin-bottom:22px;line-height:1.6}
.exam-lic-input{width:100%;padding:12px 14px;border:2px solid #dcdcdc;border-radius:8px;font-size:16px;font-family:'Courier New',monospace;text-align:center;letter-spacing:2px;margin-bottom:12px;text-transform:uppercase;box-sizing:border-box}
.exam-lic-input:focus{outline:none;border-color:#1a237e}
.exam-lic-btn{width:100%;padding:12px;background:#1a237e;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;margin-bottom:8px}
.exam-lic-btn:hover{background:#12185c}
.exam-lic-err{font-size:12px;color:#c0392b;min-height:18px;margin-bottom:6px}
.exam-lic-contact{font-size:12px;color:#666}
</style>
<div id="examLicScreen">
  <div class="exam-lic-box">
    <div class="exam-lic-icon">&#9878;</div>
    <div class="exam-lic-title">Examination Tool</div>
    <div class="exam-lic-sub">Enter your license key to continue</div>
    <input type="text" class="exam-lic-input" id="examLicKeyInput" placeholder="EXAM-XXXX-XXXX" maxlength="15" oninput="this.value=this.value.toUpperCase()">
    <div class="exam-lic-err" id="examLicErr"></div>
    <button class="exam-lic-btn" onclick="examActivateLicense()">Activate</button>
    <div class="exam-lic-contact">For license key contact: <strong>+91 9666889919</strong></div>
  </div>
</div>
<script>
var EXAM_FB_API_KEY    = 'AIzaSyAVCpa5Pji_G9XMR0CnbcV82Mpfbp6Ou9Y';
var EXAM_FB_PROJECT_ID = 'pendency-report-7669a';
var EXAM_COLLECTION    = 'licenses_examinationtool';
var EXAM_CONTACT_PHONE = '+91 9666889919';
var EXAM_LICENSE_KEY   = '';

function examSetLicErr(msg){ document.getElementById('examLicErr').textContent = msg; }

function examLaunchApp(key){
  EXAM_LICENSE_KEY = key;
  document.getElementById('examLicScreen').style.display = 'none';
  document.getElementById('examCard').style.display = 'block';
}

function examActivateLicense(){
  var key = document.getElementById('examLicKeyInput').value.trim().toUpperCase();
  if(!key){ examSetLicErr('Please enter your license key'); return; }
  if(key.length < 10){ examSetLicErr('Invalid license key format'); return; }
  examSetLicErr('Verifying\u2026');

  var url = 'https://firestore.googleapis.com/v1/projects/'+EXAM_FB_PROJECT_ID+'/databases/(default)/documents/'+EXAM_COLLECTION+'?key='+EXAM_FB_API_KEY;
  fetch(url)
    .then(function(r){ return r.json(); })
    .then(function(data){
      if(!data.documents){ examSetLicErr('Invalid license key. Contact: '+EXAM_CONTACT_PHONE); return; }
      var found = null;
      data.documents.forEach(function(doc){
        var f = doc.fields||{};
        if(f.licenseKey && f.licenseKey.stringValue === key) found = f;
      });
      if(!found){ examSetLicErr('Invalid license key. Contact: '+EXAM_CONTACT_PHONE); return; }
      var active = found.active ? found.active.booleanValue : false;
      if(!active){ examSetLicErr('This key is inactive. Contact: '+EXAM_CONTACT_PHONE); return; }
      var credits = found.creditsRemaining ? found.creditsRemaining.integerValue : '0';
      if(parseInt(credits,10) <= 0){ examSetLicErr('No credits remaining. Contact: '+EXAM_CONTACT_PHONE+' to top up.'); return; }

      localStorage.setItem('exam_key', key);
      localStorage.setItem('exam_last', new Date().toISOString());
      examLaunchApp(key);
    })
    .catch(function(){ examSetLicErr('No internet. Please connect and try again.'); });
}

window.addEventListener('load', function(){
  var storedKey  = localStorage.getItem('exam_key');
  var storedLast = localStorage.getItem('exam_last');
  if(storedKey && storedLast){
    var diffHours = (new Date() - new Date(storedLast)) / (1000*60*60);
    if(diffHours < 8){ examLaunchApp(storedKey); return; }
    document.getElementById('examLicKeyInput').value = storedKey;
  }
});
</script>

  <div class="card" id="examCard" style="display:none">
    <h1>&#9878; Examination Tool</h1>
    <div class="subtitle">Prepares Examination of Accused / Framing of Charges from a charge sheet</div>

    <div class="notice">
      Upload the <strong>charge sheet only</strong> — not the full case bundle (witness statements,
      FIR, medical reports, etc.). Each file uploaded counts as one examination.
    </div>

    <label>Court Name</label>
    <input type="text" id="court" placeholder="e.g. Judicial Magistrate of I Class, Puttur">

    <label>Judge Name (optional, for Framing of Charges)</label>
    <input type="text" id="judge" placeholder="">

    <label>Court Type</label>
    <div class="radio-row">
      <label><input type="radio" name="courtType" value="magistrate" checked> Magistrate Court (JM / JFCM / JMFC)</label>
      <label><input type="radio" name="courtType" value="sessions"> Sessions Court (ASJ / Sessions Judge / Special Court)</label>
    </div>

    <div id="modeBlock">
      <label>Classification Mode (Magistrate only)</label>
      <div class="radio-row">
        <label><input type="radio" name="mode" value="auto" checked> Automatic (AI decides Summons or Warrant)</label>
        <label><input type="radio" name="mode" value="manual"> Manual (I will select)</label>
      </div>
    </div>

    <div id="manualBlock">
      <label>Examination Type</label>
      <div class="radio-row" style="flex-direction:column;gap:8px">
        <label><input type="radio" name="manualRef" value="251" checked> Sec.251 Cr.P.C. — IPC, Summons case</label>
        <label><input type="radio" name="manualRef" value="239"> Sec.239 Cr.P.C. — IPC, Warrant case</label>
        <label><input type="radio" name="manualRef" value="262"> Sec.262 BNSS — BNS, Summons case</label>
        <label><input type="radio" name="manualRef" value="274"> Sec.274 BNSS — BNS, Warrant case</label>
      </div>
    </div>

    <label>Charge Sheet Files (.pdf, .docx, .doc — up to 20 at once)</label>
    <input type="file" id="files" multiple accept=".pdf,.docx,.doc">

    <button id="runBtn" onclick="runProcess()">Prepare Documents</button>
    <div id="resultBox"></div>
  </div>

<script>
function toggleFields(){
  var ct = document.querySelector('input[name=courtType]:checked').value;
  document.getElementById('modeBlock').style.display = (ct === 'magistrate') ? 'block' : 'none';
  var mode = document.querySelector('input[name=mode]:checked') ? document.querySelector('input[name=mode]:checked').value : 'auto';
  document.getElementById('manualBlock').style.display = (ct === 'magistrate' && mode === 'manual') ? 'block' : 'none';
}
document.querySelectorAll('input[name=courtType]').forEach(function(r){ r.addEventListener('change', toggleFields); });
document.addEventListener('change', function(e){
  if(e.target.name === 'mode') toggleFields();
});

function runProcess(){
  var key = EXAM_LICENSE_KEY;
  var court = document.getElementById('court').value.trim();
  var judge = document.getElementById('judge').value.trim();
  var courtType = document.querySelector('input[name=courtType]:checked').value;
  var modeEl = document.querySelector('input[name=mode]:checked');
  var mode = modeEl ? modeEl.value : 'auto';
  var refEl = document.querySelector('input[name=manualRef]:checked');
  var manualRef = refEl ? refEl.value : '';
  var files = document.getElementById('files').files;

  var box = document.getElementById('resultBox');
  box.style.display = 'block'; box.className = ''; box.textContent = '';

  if(!key){ box.className='err'; box.textContent='Please enter your license key.'; return; }
  if(!court){ box.className='err'; box.textContent='Please enter the court name.'; return; }
  if(!files.length){ box.className='err'; box.textContent='Please choose at least one file.'; return; }

  var btn = document.getElementById('runBtn');
  btn.disabled = true; btn.textContent = 'Processing\u2026';
  box.textContent = 'Reading and preparing documents \u2014 this can take a little while for scanned files.';

  var fd = new FormData();
  fd.append('license_key', key);
  fd.append('court', court);
  fd.append('judge', judge);
  fd.append('court_type', courtType);
  fd.append('mode', mode);
  fd.append('manual_ref', manualRef);
  for(var i=0;i<files.length;i++) fd.append('files', files[i]);

  fetch('/api/process', { method:'POST', body: fd })
    .then(function(r){
      if(!r.ok){ return r.json().then(function(d){ throw new Error(d.error || ('Server error ' + r.status)); }); }
      return r.blob();
    })
    .then(function(blob){
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url; a.download = 'Examination documents.zip'; a.click();
      URL.revokeObjectURL(url);
      box.className = 'ok';
      box.textContent = 'Done \u2014 your documents have downloaded.';
    })
    .catch(function(e){
      box.className = 'err';
      box.textContent = 'Error: ' + e.message;
    })
    .finally(function(){
      btn.disabled = false; btn.textContent = 'Prepare Documents';
    });
}
toggleFields();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
