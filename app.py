"""
Web app: one or many barcodes → ZIP with Word + PNG per item (Code 128, A4).
Each item uses a user-chosen folder/file name inside the ZIP.
"""
import io
import re
import zipfile
import tempfile
import os
from flask import Flask, request, render_template_string, send_file
from barcode_footer import generate_barcode_image_bytes, create_document_with_barcode

app = Flask(__name__)

MAX_ITEMS = 50


def sanitize_label(name):
    """
    Safe folder / basename from user input (no extension).
    Returns None if nothing usable remains.
    """
    s = re.sub(r"[^\w\-.]", "_", str(name).strip())
    s = s.strip("._")
    if not s:
        return None
    return s[:80]


def image_filename_from_barcode(barcode_value):
    """PNG basename from encoded barcode value (not the Word doc name)."""
    s = re.sub(r"[^\w\-]", "_", str(barcode_value).strip())[:80]
    return (s or "barcode") + ".png"


def build_one_pair_zip(sanitized_label, barcode_value):
    """Single item: ZIP with one folder containing .docx + .png."""
    barcode_image_bytes = generate_barcode_image_bytes(barcode_value)
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp.write(barcode_image_bytes)
        tmp_path = tmp.name
    try:
        docx_buffer = io.BytesIO()
        create_document_with_barcode(barcode_value, docx_buffer, barcode_image_path=tmp_path)
        docx_bytes = docx_buffer.getvalue()
    finally:
        os.unlink(tmp_path)

    zip_buffer = io.BytesIO()
    folder = sanitized_label
    png_name = image_filename_from_barcode(barcode_value)
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"{folder}/{folder}.docx", docx_bytes)
        zf.writestr(f"{folder}/{png_name}", barcode_image_bytes)
    zip_buffer.seek(0)
    return zip_buffer


def build_multi_zip(pairs):
    """pairs: list of (sanitized_label, barcode_value)."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for sanitized_label, barcode_value in pairs:
            barcode_image_bytes = generate_barcode_image_bytes(barcode_value)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                tmp.write(barcode_image_bytes)
                tmp_path = tmp.name
            try:
                docx_buffer = io.BytesIO()
                create_document_with_barcode(
                    barcode_value, docx_buffer, barcode_image_path=tmp_path
                )
                docx_bytes = docx_buffer.getvalue()
            finally:
                os.unlink(tmp_path)
            folder = sanitized_label
            png_name = image_filename_from_barcode(barcode_value)
            zf.writestr(f"{folder}/{folder}.docx", docx_bytes)
            zf.writestr(f"{folder}/{png_name}", barcode_image_bytes)
    zip_buffer.seek(0)
    return zip_buffer


HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Word Barcode Generator - Code 128</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'DM Sans', system-ui, -apple-system, sans-serif;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 1.5rem;
      background: linear-gradient(145deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
      color: #e2e8f0;
    }
    .wrap {
      width: 100%;
      max-width: 560px;
    }
    .card {
      background: rgba(30, 41, 59, 0.8);
      backdrop-filter: blur(12px);
      padding: 2.25rem;
      border-radius: 20px;
      border: 1px solid rgba(71, 85, 105, 0.5);
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.4);
    }
    .card-header {
      text-align: center;
      margin-bottom: 1.5rem;
    }
    .icon {
      width: 48px;
      height: 48px;
      margin: 0 auto 1rem;
      background: linear-gradient(135deg, #3b82f6, #8b5cf6);
      border-radius: 14px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 1.5rem;
    }
    h1 {
      font-size: 1.5rem;
      font-weight: 700;
      color: #f8fafc;
      letter-spacing: -0.02em;
      margin-bottom: 0.35rem;
    }
    .sub {
      color: #94a3b8;
      font-size: 0.9rem;
      font-weight: 500;
    }
    .hint {
      color: #64748b;
      font-size: 0.8rem;
      margin-top: 0.35rem;
      line-height: 1.4;
    }
    form { margin-top: 1.25rem; }
    label {
      display: block;
      font-size: 0.8rem;
      font-weight: 600;
      color: #cbd5e1;
      margin-bottom: 0.35rem;
    }
    input[type="text"] {
      width: 100%;
      padding: 0.65rem 0.85rem;
      font-size: 0.95rem;
      font-family: inherit;
      color: #f8fafc;
      background: rgba(15, 23, 42, 0.6);
      border: 1px solid rgba(71, 85, 105, 0.6);
      border-radius: 10px;
      transition: border-color 0.2s, box-shadow 0.2s;
    }
    input[type="text"]::placeholder { color: #64748b; }
    input[type="text"]:focus {
      outline: none;
      border-color: #3b82f6;
      box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.25);
    }
    .row-block {
      padding: 1rem;
      margin-bottom: 1rem;
      background: rgba(15, 23, 42, 0.45);
      border-radius: 14px;
      border: 1px solid rgba(71, 85, 105, 0.35);
      position: relative;
    }
    .row-block .row-title {
      font-size: 0.75rem;
      font-weight: 700;
      color: #94a3b8;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      margin-bottom: 0.75rem;
    }
    .field { margin-bottom: 0.75rem; }
    .field:last-of-type { margin-bottom: 0; }
    .row-actions {
      display: flex;
      justify-content: flex-end;
      margin-top: 0.5rem;
    }
    .btn-remove {
      padding: 0.35rem 0.65rem;
      font-size: 0.8rem;
      font-weight: 600;
      font-family: inherit;
      color: #f87171;
      background: transparent;
      border: 1px solid rgba(248, 113, 113, 0.35);
      border-radius: 8px;
      cursor: pointer;
    }
    .btn-remove:hover { background: rgba(248, 113, 113, 0.1); }
    .btn-remove:disabled { opacity: 0.35; cursor: not-allowed; }
    .btn-add {
      width: 100%;
      padding: 0.65rem 1rem;
      font-size: 0.9rem;
      font-weight: 600;
      font-family: inherit;
      color: #93c5fd;
      background: rgba(59, 130, 246, 0.12);
      border: 1px dashed rgba(59, 130, 246, 0.45);
      border-radius: 12px;
      cursor: pointer;
      margin-bottom: 1rem;
    }
    .btn-add:hover { background: rgba(59, 130, 246, 0.2); }
    .btn-add:disabled { opacity: 0.5; cursor: not-allowed; }
    .btn {
      width: 100%;
      padding: 0.9rem 1.25rem;
      font-size: 1rem;
      font-weight: 600;
      font-family: inherit;
      color: #fff;
      background: linear-gradient(135deg, #3b82f6, #2563eb);
      border: none;
      border-radius: 12px;
      cursor: pointer;
      transition: transform 0.15s, box-shadow 0.2s;
      box-shadow: 0 4px 14px rgba(59, 130, 246, 0.4);
    }
    .btn:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(59, 130, 246, 0.45); }
    .btn:active { transform: translateY(0); }
    .btn:disabled { opacity: 0.7; cursor: not-allowed; transform: none; }
    .error {
      color: #f87171;
      font-size: 0.875rem;
      margin-top: 0.75rem;
      padding: 0.5rem 0;
    }
    .credit {
      margin-top: 2rem;
      text-align: center;
      font-size: 0.85rem;
      color: #64748b;
    }
    .credit strong { color: #94a3b8; font-weight: 600; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <div class="card-header">
        <div class="icon">📄</div>
        <h1>Word Barcode Generator</h1>
        <p class="sub">Code 128 · A4 · Word + barcode image (ZIP)</p>
        <p class="hint">Name each set (folder in the ZIP). The Word file uses that name; the PNG is named from the barcode value.</p>
      </div>
      <form method="post" action="/generate" id="f">
        <div id="rows"></div>
        <button type="button" class="btn-add" id="addRow">+ Add another file</button>
        <button type="submit" class="btn" id="btn">Generate &amp; download ZIP</button>
      </form>
      {% if error %}<p class="error">{{ error }}</p>{% endif %}
    </div>
    <p class="credit">Built by <strong>Brian Kyle L. Salor</strong></p>
  </div>
  <script>
    var MAX = """ + str(MAX_ITEMS) + """;
    var rowCount = 0;

    function rowHtml(n) {
      return (
        '<div class="row-block" data-row="' + n + '">' +
        '<div class="row-title">Set #' + n + '</div>' +
        '<div class="field">' +
        '<label>Folder &amp; file name <span style="font-weight:400;color:#64748b">(no extension)</span></label>' +
        '<input type="text" name="file_label" required placeholder="e.g. L1_Invoice or SKU_Box12" maxlength="100">' +
        '</div>' +
        '<div class="field">' +
        '<label>Barcode value</label>' +
        '<input type="text" name="barcode_value" required placeholder="e.g. 123456789">' +
        '</div>' +
        '<div class="row-actions">' +
        '<button type="button" class="btn-remove" data-remove>Remove</button>' +
        '</div>' +
        '</div>'
      );
    }

    function refreshRemoveState() {
      var blocks = document.querySelectorAll('.row-block');
      var canRemove = blocks.length > 1;
      blocks.forEach(function(b, i) {
        var title = b.querySelector('.row-title');
        if (title) { title.textContent = 'Set #' + (i + 1); }
        var btn = b.querySelector('[data-remove]');
        if (btn) { btn.disabled = !canRemove; }
      });
      document.getElementById('addRow').disabled = blocks.length >= MAX;
    }

    function addRow() {
      if (rowCount >= MAX) return;
      rowCount++;
      var wrap = document.getElementById('rows');
      var div = document.createElement('div');
      div.innerHTML = rowHtml(rowCount);
      wrap.appendChild(div.firstElementChild);
      refreshRemoveState();
    }

    document.getElementById('rows').addEventListener('click', function(e) {
      if (e.target.matches('[data-remove]') && !e.target.disabled) {
        e.target.closest('.row-block').remove();
        refreshRemoveState();
      }
    });

    document.getElementById('addRow').addEventListener('click', addRow);

    for (var i = 0; i < 1; i++) { addRow(); }

    document.getElementById('f').onsubmit = function() {
      document.getElementById('btn').disabled = true;
      document.getElementById('btn').textContent = 'Generating…';
    };
  </script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML)


def _parse_pairs_from_form():
    """Return (pairs, error_message) where pairs is list of (sanitized_label, value) or None."""
    labels = request.form.getlist("file_label")
    values = request.form.getlist("barcode_value")
    if len(labels) != len(values):
        return None, "Form mismatch. Please refresh and try again."

    raw_rows = []
    for i, (raw_label, raw_val) in enumerate(zip(labels, values)):
        label = (raw_label or "").strip()
        value = (raw_val or "").strip()
        if not label and not value:
            continue
        if not label:
            return None, f"Set #{i + 1}: enter a folder and file name."
        if not value:
            return None, f"Set #{i + 1}: enter a barcode value."
        safe = sanitize_label(label)
        if not safe:
            return None, f"Set #{i + 1}: use letters, numbers, dashes, or dots in the name."
        raw_rows.append((safe, value, label))

    if not raw_rows:
        return None, "Add at least one set with a name and barcode value."

    if len(raw_rows) > MAX_ITEMS:
        return None, f"Maximum {MAX_ITEMS} sets per download."

    seen = {}
    pairs = []
    for safe, value, original in raw_rows:
        if safe in seen:
            return (
                None,
                f"Duplicate name “{safe}” after cleaning (sets “{seen[safe]}” and “{original}”). "
                "Use unique names.",
            )
        seen[safe] = original
        pairs.append((safe, value))

    return pairs, None


@app.route("/generate", methods=["POST"])
def generate():
    pairs, err = _parse_pairs_from_form()
    if err:
        return render_template_string(HTML, error=err), 400
    try:
        if len(pairs) == 1:
            label, value = pairs[0]
            zip_buffer = build_one_pair_zip(label, value)
            zip_name = f"{label}.zip"
        else:
            zip_buffer = build_multi_zip(pairs)
            zip_name = "barcodes_batch.zip"
    except Exception as e:
        return render_template_string(HTML, error=str(e)), 500
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip",
    )


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
