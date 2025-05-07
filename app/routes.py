import os
from flask import Blueprint, render_template, request, send_from_directory, current_app, redirect, url_for, flash
import pandas as pd
from .utils import process_clock_times

main_bp = Blueprint('main', __name__)

@main_bp.route('/')
def index():
    return render_template('index.html')

@main_bp.route('/upload', methods=['POST'])
def upload_file():
    if 'clock_ins_file' not in request.files:
        flash('No file selected', 'error')
        return redirect(url_for('main.index'))
    
    clock_ins_file = request.files['clock_ins_file']
    if clock_ins_file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('main.index'))
    
    # Save the file
    filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], clock_ins_file.filename)
    clock_ins_file.save(filepath)
    
    return redirect(url_for('main.process_file', filename=clock_ins_file.filename))

@main_bp.route('/process/<filename>', methods=['GET', 'POST'])
def process_file(filename):
    if request.method == 'GET':
        return render_template('process.html', filename=filename)
    
    try:
        # Get rounding interval
        interval = int(request.form.get('interval', 15))
        
        # Read file from specific sheet
        filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], filename)
        df = pd.read_excel(filepath, sheet_name='9. Payroll')
        
        # Process times
        df_processed = process_clock_times(df, interval)
        
        # Save processed file with adjusted columns
        output_filename = f'processed_{filename}'
        output_path = os.path.join(current_app.config['PROCESSED_FOLDER'], output_filename)
        
        # Create ExcelWriter
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write DataFrame
            df_processed.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # Get worksheet
            worksheet = writer.sheets['Sheet1']
            
            # Adjust column widths
            for idx, col in enumerate(df_processed.columns):
                # Get maximum content length
                max_length = max(
                    df_processed[col].astype(str).apply(len).max(),  # content length
                    len(str(col))  # column name length
                )
                # Adjust column width (add some extra space)
                worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2
        
        return redirect(url_for('main.download_file', filename=output_filename))
    except Exception as e:
        flash(f'Error processing file: {str(e)}', 'error')
        return redirect(url_for('main.index'))

@main_bp.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(current_app.config['PROCESSED_FOLDER'], filename, as_attachment=True) 