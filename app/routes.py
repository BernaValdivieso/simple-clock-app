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
        # Obtener el intervalo de redondeo, formato de hora y decimales
        interval = int(request.form.get('interval', 15))
        time_format = request.form.get('time_format', '24')
        decimals = request.form.get('decimals', 'all')
        
        # Leer el archivo de la hoja específica
        filepath = os.path.join(current_app.config['UPLOAD_FOLDER'], filename)
        df = pd.read_excel(filepath, sheet_name='9. Payroll')
        
        # Procesar los tiempos
        df_processed = process_clock_times(df, interval, decimals, time_format)
        
        # Guardar el archivo procesado con columnas ajustadas
        output_filename = f'processed_{filename}'
        output_path = os.path.join(current_app.config['PROCESSED_FOLDER'], output_filename)
        
        # Crear un ExcelWriter
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Escribir el DataFrame
            df_processed.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # Obtener la hoja de trabajo
            worksheet = writer.sheets['Sheet1']
            
            # Agregar filtros a todas las columnas
            worksheet.auto_filter.ref = worksheet.dimensions
            
            # Ajustar el ancho de las columnas
            for idx, col in enumerate(df_processed.columns):
                # Obtener la longitud máxima del contenido de la columna
                max_length = max(
                    df_processed[col].astype(str).apply(len).max(),  # longitud del contenido
                    len(str(col))  # longitud del nombre de la columna
                )
                # Ajustar el ancho de la columna (agregar un poco de espacio extra)
                worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2
        
        return redirect(url_for('main.download_file', filename=output_filename))
    except Exception as e:
        flash(f'Error processing file: {str(e)}', 'error')
        return redirect(url_for('main.index'))

@main_bp.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(current_app.config['PROCESSED_FOLDER'], filename, as_attachment=True) 