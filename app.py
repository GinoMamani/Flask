import os
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import secrets
import pandas as pd

app = Flask(__name__, template_folder='Templates')
app.secret_key = secrets.token_hex(24)

# Archivo Excel para almacenar registros
excel_file = 'Solicitud de repuestos y Herramientas.xlsx'

# Simulated database
database = []

# Home page displaying all records
@app.route('/')
def index():
    return render_template('index.html', records=database)

# Form for creating a new record
@app.route('/new', methods=['GET', 'POST'])
def new():
    if request.method == 'POST':
        personal = request.form['personal']
        herramienta = request.form['herramienta']
        cantidad = int(request.form['cantidad'])
        prioridad = request.form['prioridad']
        comentario = request.form.get('comentario', '')
        link = request.form['link']

        if personal and herramienta and cantidad and prioridad and link:
            new_record = {
                'personal': personal,
                'herramienta': herramienta,
                'cantidad': cantidad,
                'prioridad': prioridad,
                'comentario': comentario,
                'link': link
            }
            database.append(new_record)
            flash('Registro creado exitosamente!', 'success')

            # Guardar registros en el archivo Excel
            save_records_to_excel()

            # Llamar a la función send_email después de guardar el registro
            send_email(new_record)

            return redirect(url_for('index'))
        else:
            flash('Todos los campos son requeridos excepto el comentario.', 'danger')

    return render_template('form.html')

def save_records_to_excel():
    # Crear un DataFrame de pandas a partir de la base de datos simulada
    df = pd.DataFrame(database)
    
    if df.empty:
        # Si el DataFrame está vacío, añadir encabezados
        df = pd.DataFrame(columns=['personal', 'herramienta', 'cantidad', 'prioridad', 'comentario', 'link'])
    
    # Guardar el DataFrame a un archivo Excel
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)

# Definir las variables directamente en el código
OUTLOOK_EMAIL = 'Solicitudrepuestos@outlook.com'
OUTLOOK_PASSWORD = 'xperiaplay13637'
CONTRASEÑA_APLICACION = 'rsgjsedavsrnpxwb'

# Editing a record
@app.route('/edit/<int:record_id>', methods=['GET', 'POST'])
def edit(record_id):
    record = database[record_id]
    if request.method == 'POST':
        record['personal'] = request.form['personal']
        record['herramienta'] = request.form['herramienta']
        record['cantidad'] = int(request.form['cantidad'])
        record['prioridad'] = request.form['prioridad']
        record['comentario'] = request.form.get('comentario', '')
        record['link'] = request.form['link']

        flash('Registro actualizado exitosamente!', 'success')

        # Guardar registros actualizados en el archivo Excel
        save_records_to_excel()

        return redirect(url_for('index'))

    return render_template('form.html', record=record, record_id=record_id)

# Deleting a record
@app.route('/delete/<int:record_id>')
def delete(record_id):
    del database[record_id]
    flash('Registro eliminado exitosamente!', 'success')

    # Guardar registros actualizados en el archivo Excel
    save_records_to_excel()

    return redirect(url_for('index'))

# Sending an email
def send_email(record):
    recipient = ['leonardo.daviran@overall.com.pe', 'mantenimientoelectronico@overall.com.pe','gino.mamani@overall.com.pe']
    subject = f"SOLICITUD DE HERRAMIENTA - {record['cantidad']} {record['herramienta']} - Prioridad: {record['prioridad']}"
    body = f"El personal {record['personal']} solicita {record['cantidad']} {record['herramienta']}.\n\nEn el siguiente link puede ver lo solicitado: {record['link']}\n\n{record['comentario']}"

    msg = MIMEMultipart()
    msg['From'] = OUTLOOK_EMAIL
    msg['To'] = ', '.join(recipient)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP('smtp-mail.outlook.com', 587) as server:
            server.starttls()
            server.login(OUTLOOK_EMAIL, CONTRASEÑA_APLICACION)
            server.sendmail(OUTLOOK_EMAIL, recipient, msg.as_string())
        flash('Correo electrónico enviado exitosamente!', 'success')
    except Exception as e:
        print(f'Error: {str(e)}')  # Añadir esta línea para imprimir el error
        flash(f'No se pudo enviar el correo electrónico: {str(e)}', 'danger')

# Ruta para exportar la tabla a Excel
@app.route('/export')
def export():
    # Asegurar que los registros estén guardados en el archivo Excel
    save_records_to_excel()
    # Enviar el archivo Excel al navegador
    return send_file(excel_file, download_name='Solicitud de repuestos y Herramientas.xlsx', as_attachment=True)

# Run the app
if __name__ == '__main__':
    app.run
