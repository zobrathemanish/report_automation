import os
import time
import re
import pandas as pd
from threading import Thread
from flask import Flask, request, redirect, url_for, jsonify, send_from_directory, render_template_string
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_COLOR_INDEX

app = Flask(__name__)

# Configure upload and output directories.
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Global dictionary to store progress information.
progress_data = {
    'progress': 0,
    'output_filename': None
}

# Helper function to capture formatting from a run.
def get_run_formatting(run):
    fmt = {
        'bold': run.font.bold,
        'italic': run.font.italic,
        'underline': run.font.underline,
        'size': run.font.size,
        'name': run.font.name,
        'color': run.font.color.rgb if run.font.color and run.font.color.rgb else None,
        # You can add more formatting properties if needed.
    }
    return fmt

# Helper function to apply formatting to a target run.
def apply_formatting(target_run, fmt):
    target_run.font.bold = fmt.get('bold')
    target_run.font.italic = fmt.get('italic')
    target_run.font.underline = fmt.get('underline')
    target_run.font.size = fmt.get('size')
    target_run.font.name = fmt.get('name')
    if fmt.get('color'):
        target_run.font.color.rgb = fmt.get('color')

# --------------------------
# Processing Function
# --------------------------
def process_files(excel_path, word_path, highlight_flag):
    """
    Follows the notebook structure:
      1. Load Excel data.
      2. Load the Word template.
      3. Replace placeholders while preserving run formatting.
      4. Add notifications if there are mismatches.
      5. Save the generated report.
    """
    try:
        # --------------------------
        # 1. Load Excel Data
        # --------------------------
        progress_data['progress'] = 10
        df = pd.read_excel(excel_path, sheet_name=1)  # Read second sheet
        df.rename(columns=lambda x: x.strip(), inplace=True)

        # Check for required columns.
        missing_columns = []
        if "Metric" not in df.columns:
            missing_columns.append("Metric")
        if "Value" not in df.columns:
            missing_columns.append("Value")
            
        if missing_columns:
            print("Missing required column(s):", missing_columns)
            stats_dict = {}  # Proceed with an empty dictionary
        else:
            stats_dict = df.set_index('Metric')['Value'].astype(str).to_dict()
            # Convert both keys and values to strings before stripping.
            stats_dict = {str(key).strip(): str(value).strip() for key, value in stats_dict.items()}

        print("✅ Cleaned Dictionary Keys:", stats_dict.keys())
        time.sleep(1)
        progress_data['progress'] = 30

        # --------------------------
        # 2. Load the Word Template
        # --------------------------
        doc = Document(word_path)
        time.sleep(1)
        progress_data['progress'] = 40

        # --------------------------
        # 3. Extract Placeholders (for diagnostics)
        # --------------------------
        def extract_placeholders(document):
            placeholders = set()
            pattern = re.compile(r"{{\s*(.*?)\s*}}")
            for para in document.paragraphs:
                matches = pattern.findall(para.text)
                placeholders.update(matches)
            return placeholders

        doc_placeholders = extract_placeholders(doc)
        print("🔍 Placeholders in Template:", doc_placeholders)
        missing_in_excel = doc_placeholders - set(stats_dict.keys())
        if missing_in_excel:
            print("⚠️ The following placeholders exist in the template but NOT in Excel:")
            print(missing_in_excel)
        missing_in_doc = set(stats_dict.keys()) - doc_placeholders
        if missing_in_doc:
            print("⚠️ The following Excel keys exist but are NOT used in the template:")
            print(missing_in_doc)

        # --------------------------
        # 4. Replacement Function (Preserving Run Formatting)
        # --------------------------
        def replace_placeholders_in_docx(document, replacements, apply_highlight=False):
            """
            Processes each paragraph and run. For each run, splits its text around placeholders of
            the form {{ key }}. For each text segment, a new run is created that inherits the original
            run’s formatting (bold, italic, size, color, etc.). If a segment is a replacement and
            highlighting is enabled, the run is highlighted.
            """
            pattern = re.compile(r"{{\s*(.*?)\s*}}")
            for para in document.paragraphs:
                new_runs = []  # List of tuples: (text, is_replaced, formatting dict)
                for run in para.runs:
                    fmt = get_run_formatting(run)
                    text = run.text
                    last_index = 0
                    if not pattern.search(text):
                        new_runs.append((text, False, fmt))
                    else:
                        for match in pattern.finditer(text):
                            start, end = match.span()
                            key = match.group(1).strip()
                            # Append text before the placeholder.
                            if start > last_index:
                                new_runs.append((text[last_index:start], False, fmt))
                            # Append the replacement value if found.
                            if key in replacements:
                                new_runs.append((replacements[key], True, fmt))
                            else:
                                new_runs.append((match.group(0), False, fmt))
                            last_index = end
                        # Append any remaining text.
                        if last_index < len(text):
                            new_runs.append((text[last_index:], False, fmt))
                # Remove all existing runs.
                for run in para.runs:
                    run.text = ""
                # Rebuild the paragraph with new runs, applying formatting.
                for seg_text, is_replaced, fmt in new_runs:
                    new_run = para.add_run(seg_text)
                    apply_formatting(new_run, fmt)
                    if is_replaced and apply_highlight:
                        new_run.font.highlight_color = WD_COLOR_INDEX.YELLOW

        replace_placeholders_in_docx(doc, stats_dict, apply_highlight=highlight_flag)

        # --------------------------
        # 5. Add Notification Paragraphs (if mismatches exist)
        # --------------------------
        if missing_in_excel:
            note = ("Note: The following placeholders were not found in the Excel file: " +
                    ", ".join(missing_in_excel))
            doc.add_paragraph(note)
        if missing_in_doc:
            note2 = ("Note: The following keys from the Excel file were not used in the template: " +
                     ", ".join(missing_in_doc))
            doc.add_paragraph(note2)

        time.sleep(1)
        progress_data['progress'] = 80

        # --------------------------
        # 6. Save the Generated Report
        # --------------------------
        output_filename = "generated_report.docx"
        output_filepath = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        doc.save(output_filepath)
        time.sleep(1)
        progress_data['progress'] = 100
        progress_data['output_filename'] = output_filename
        print(f"✅ Report generated successfully: {output_filepath}")

    except Exception as e:
        progress_data['progress'] = 100
        print("Error processing files:", e)


# -------------------------
# Flask Routes
# -------------------------

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    """Upload form to provide an Excel file and a Word template. Optionally enable highlighting."""
    if request.method == 'POST':
        if 'excel_file' not in request.files or 'word_file' not in request.files:
            return "Missing file(s)", 400

        excel_file = request.files['excel_file']
        word_file = request.files['word_file']

        if excel_file.filename == '' or word_file.filename == '':
            return "No selected file(s)", 400

        # Save the uploaded files.
        excel_filename = secure_filename(excel_file.filename)
        word_filename = secure_filename(word_file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        word_path = os.path.join(app.config['UPLOAD_FOLDER'], word_filename)
        excel_file.save(excel_path)
        word_file.save(word_path)

        # Determine if highlighting is enabled.
        highlight_flag = request.form.get("highlight") in ["true", "on"]

        # Reset progress information.
        progress_data['progress'] = 0
        progress_data['output_filename'] = None

        # Start processing in a background thread.
        thread = Thread(target=process_files, args=(excel_path, word_path, highlight_flag))
        thread.start()

        return redirect(url_for('progress_page'))

    # GET: Render the upload form.
    return render_template_string('''
    <!doctype html>
    <html lang="en">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Upload Files</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      </head>
      <body>
        <div class="container mt-5">
          <h1 class="mb-4">Upload Excel Data and Word Template</h1>
          <form method="post" enctype="multipart/form-data">
            <div class="mb-3">
              <label for="excel_file" class="form-label">Excel File (.xlsx, .xls)</label>
              <input class="form-control" type="file" name="excel_file" id="excel_file" accept=".xlsx,.xls">
            </div>
            <div class="mb-3">
              <label for="word_file" class="form-label">Word Template (.docx)</label>
              <input class="form-control" type="file" name="word_file" id="word_file" accept=".docx">
            </div>
            <div class="form-check mb-3">
              <input class="form-check-input" type="checkbox" value="true" id="highlight" name="highlight">
              <label class="form-check-label" for="highlight">
                Highlight replaced values
              </label>
            </div>
            <button type="submit" class="btn btn-primary">Submit</button>
          </form>
        </div>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
      </body>
    </html>
    ''')

@app.route('/progress_page')
def progress_page():
    """Display a progress bar and show a download button when processing is complete."""
    return render_template_string('''
    <!doctype html>
    <html lang="en">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Processing...</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
      </head>
      <body>
        <div class="container mt-5">
          <h1>Processing...</h1>
          <div class="progress">
            <div id="progressBar" class="progress-bar progress-bar-striped" role="progressbar" style="width: 0%;" aria-valuenow="0" aria-valuemin="0" aria-valuemax="100">0%</div>
          </div>
          <div id="download" class="mt-3" style="display:none;">
            <a href="/download" class="btn btn-success">Download Generated Report</a>
          </div>
        </div>
        <script>
          function checkProgress() {
            $.getJSON("/progress", function(data) {
              var progress = data.progress;
              $("#progressBar").css("width", progress + "%");
              $("#progressBar").attr("aria-valuenow", progress);
              $("#progressBar").text(progress + "%");
              if (progress >= 100) {
                $("#download").show();
              } else {
                setTimeout(checkProgress, 1000);
              }
            });
          }
          $(document).ready(function(){
            checkProgress();
          });
        </script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
      </body>
    </html>
    ''')

@app.route('/progress')
def progress():
    """Return the current processing progress."""
    return jsonify(progress=progress_data['progress'])

@app.route('/download')
def download():
    """Send the generated file to the user."""
    if progress_data['output_filename']:
        return send_from_directory(app.config['OUTPUT_FOLDER'],
                                   progress_data['output_filename'],
                                   as_attachment=True)
    else:
        return "File not available", 404

if __name__ == '__main__':
    app.run(debug=True)
