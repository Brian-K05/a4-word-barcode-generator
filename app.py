"""
Web app: enter barcode value → Generate → Download Word file (Code 128, A4, barcode in footer).
"""
import io
import re
from flask import Flask, request, render_template_string, send_file
from barcode_footer import generate_word_bytes

app = Flask(__name__)


def safe_filename(value):
    """Safe filename from barcode value."""
    safe = re.sub(r'[^\w\-]', '_', str(value))[:50]
    return f"barcode_{safe}.docx"


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
      max-width: 400px;
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
      margin-bottom: 1.75rem;
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
    form { margin-top: 1.5rem; }
    label {
      display: block;
      font-size: 0.875rem;
      font-weight: 600;
      color: #cbd5e1;
      margin-bottom: 0.5rem;
    }
    input[type="text"] {
      width: 100%;
      padding: 0.85rem 1rem;
      font-size: 1rem;
      font-family: inherit;
      color: #f8fafc;
      background: rgba(15, 23, 42, 0.6);
      border: 1px solid rgba(71, 85, 105, 0.6);
      border-radius: 12px;
      margin-bottom: 1.25rem;
      transition: border-color 0.2s, box-shadow 0.2s;
    }
    input[type="text"]::placeholder { color: #64748b; }
    input[type="text"]:focus {
      outline: none;
      border-color: #3b82f6;
      box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.25);
    }
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
        <p class="sub">Code 128 · A4 · Barcode in footer</p>
      </div>
      <form method="post" action="/generate" id="f">
        <label for="value">Barcode value</label>
        <input type="text" id="value" name="barcode_value" placeholder="e.g. 123456789" required autofocus>
        <button type="submit" class="btn" id="btn">Generate & Download</button>
      </form>
      {% if error %}<p class="error">{{ error }}</p>{% endif %}
    </div>
    <p class="credit">Built by <strong>Brian Kyle L. Salor</strong></p>
  </div>
  <script>
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


@app.route("/generate", methods=["POST"])
def generate():
    value = (request.form.get("barcode_value") or "").strip()
    if not value:
        return render_template_string(HTML, error="Please enter a barcode value."), 400
    try:
        docx_bytes = generate_word_bytes(value)
    except Exception as e:
        return render_template_string(HTML, error=str(e)), 500
    return send_file(
        io.BytesIO(docx_bytes),
        as_attachment=True,
        download_name=safe_filename(value),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
