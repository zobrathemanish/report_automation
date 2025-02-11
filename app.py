import os
import time
import re
import pandas as pd
from threading import Thread
from flask import Flask, request, redirect, url_for, jsonify, send_from_directory, render_template_string
from werkzeug.utils import secure_filename
from docx import Document
from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt

app = Flask(__name__)

# Set up directories for uploads and outputs.
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Global dictionary to store progress information.
progress_data = {
    'progress': 0,
    'output_filename': None
}

def process_files(excel_filepath, word_filepath, highlight_flag):
    """
    Processes the two files:
      1. Reads the Excel file and builds a dictionary of placeholder values.
      2. Loads the Word template.
      3. Replaces placeholders with their values. If highlight_flag is True, only the replaced text is highlighted.
      4. Sets paragraph-level formatting:
         - Alignment: Justified
         - Space Before: 0 pt
         - Space After: 12 pt
         - Line Spacing: Exactly 16 pt
      5. Saves the output file.
    """
    try:
        # --- Step 1: Load Excel Data ---
        progress_data['progress'] = 10
        df = pd.read_excel(excel_filepath, sheet_name=1)
        df.rename(columns=lambda x: x.strip(), inplace=True)
        if "Metric" not in df.columns or "Value" not in df.columns:
            progress_data['progress'] = 100
            return
        # Build dictionary mapping each Metric to its Value (as a stripped string)
        stats_dict = df.set_index('Metric')['Value'].astype(str).to_dict()
        stats_dict = {key.strip(): value.strip() for key, value in stats_dict.items()}
        time.sleep(1)  # Simulate processing delay.
        progress_data['progress'] = 30

        # --- Step 2: Load Word Template ---
        doc = Document(word_filepath)
        time.sleep(1)
        progress_data['progress'] = 40

        # --- Step 3: Replace placeholders and set paragraph formatting ---
        def replace_placeholders_in_docx(doc, replacements, highlight_replacements=False):
            """
            Replaces placeholders of the form {{ key }} with corresponding values.
            After rebuilding each paragraph, sets:
              - Justified alignment
              - 0 pt before, 12 pt after spacing
              - Exactly 16 pt line spacing
            """
            placeholder_pattern = re.compile(r"{{\s*(.*?)\s*}}")
            for para in doc.paragraphs:
                original_text = para.text
                # If no placeholders found, still apply the paragraph formatting.
                if not placeholder_pattern.search(original_text):
                    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after = Pt(12)
                    para.paragraph_format.line_spacing = Pt(16)
                    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                    continue

                # Remove all runs in the paragraph.
                p_xml = para._element
                for child in list(p_xml):
                    p_xml.remove(child)

                current_index = 0
                # Process each placeholder occurrence.
                for match in placeholder_pattern.finditer(original_text):
                    start, end = match.span()
                    # Add text before the placeholder as a normal run.
                    if start > current_index:
                        pre_text = original_text[current_index:start]
                        para.add_run(pre_text)
                    key = match.group(1).strip()
                    if key in replacements:
                        replacement_value = replacements[key]
                        new_run = para.add_run(replacement_value)
                        if highlight_replacements:
                            new_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    else:
                        # No replacement found; add the placeholder back.
                        para.add_run(match.group(0))
                    current_index = end
                # Add any remaining text after the last placeholder.
                if current_index < len(original_text):
                    para.add_run(original_text[current_index:])

                # Now apply the desired paragraph formatting.
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(12)
                para.paragraph_format.line_spacing = Pt(16)
                para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY

        replace_placeholders_in_docx(doc, stats_dict, highlight_flag)
        time.sleep(1)
        progress_data['progress'] = 80

        # --- Step 4: Save the generated report ---
        output_filename = "generated_report.docx"
        output_filepath = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        doc.save(output_filepath)
        time.sleep(1)
        progress_data['progress'] = 100
        progress_data['output_filename'] = output_filename

    except Exception as e:
        progress_data['progress'] = 100
        print("Error processing files:", e)

# -------------------------
# Flask Routes
# -------------------------

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    """Upload form with a checkbox to indicate whether to highlight replaced text."""
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

        # Determine if the checkbox for highlighting is checked.
        highlight_flag = request.form.get("highlight") in ["true", "on"]

        # Reset progress info.
        progress_data['progress'] = 0
        progress_data['output_filename'] = None

        # Start processing in a background thread.
        thread = Thread(target=process_files, args=(excel_path, word_path, highlight_flag))
        thread.start()

        return redirect(url_for('progress_page'))

    # GET: Render the upload form with Bootstrap styling.
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
              <label for="excel_file" class="form-label">Current Johnson Data Excel File</label>
              <input class="form-control" type="file" name="excel_file" id="excel_file" accept=".xlsx,.xls">
            </div>
            <div class="mb-3">
              <label for="word_file" class="form-label">Word Template</label>
              <input class="form-control" type="file" name="word_file" id="word_file" accept=".docx">
            </div>
            <div class="form-check mb-3">
              <input class="form-check-input" type="checkbox" value="true" id="highlight" name="highlight">
              <label class="form-check-label" for="highlight">
                Highlight Replaced Values
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
    """Display a progress bar and, when complete, a download button."""
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
    """Send the generated output file to the user."""
    if progress_data['output_filename']:
        return send_from_directory(app.config['OUTPUT_FOLDER'],
                                   progress_data['output_filename'],
                                   as_attachment=True)
    else:
        return "File not available", 404

if __name__ == '__main__':
    app.run(debug=True)
