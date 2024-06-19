import os
from flask import Flask, render_template, request, redirect, url_for, flash
import smtplib
from email.mime.text import MIMEText
import secrets

app = Flask(__name__, template_folder='Templates')
app.secret_key = secrets.token_hex(24)

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

            # Llamar a la función send_email después de guardar el registro
            send_email(new_record)  # Esto enviará el correo al destinatario definido

            return redirect(url_for('index'))
        else:
            flash('Todos los campos son requeridos excepto el comentario.', 'danger')

    return render_template('form.html')

# Definir las variables directamente en el código
OUTLOOK_EMAIL = 'Solicitudrepuestos@outlook.com'
OUTLOOK_PASSWORD = 'xperiaplay13637'
CONTRASEÑA_APLICACION = 'mksjklbizhpccnpj'

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
        return redirect(url_for('index'))

    return render_template('form.html', record=record, record_id=record_id)

# Deleting a record
@app.route('/delete/<int:record_id>')
def delete(record_id):
    del database[record_id]
    flash('Registro eliminado exitosamente!', 'success')
    return redirect(url_for('index'))

# Sending an email
def send_email(record):
    recipient = ['leonardo.daviran@overall.com.pe', 'mantenimientoelectronico@overall.com.pe']
    subject = f"SOLICITUD DE HERRAMIENTA - {record['cantidad']} {record['herramienta']} - Prioridad: {record['prioridad']}"
    body = f"El personal {record['personal']} solicita {record['cantidad']} {record['herramienta']}.\n\nEn el siguiente link puede ver lo solicitado: {record['link']}\n\n{record['comentario']}"

    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = OUTLOOK_EMAIL
    msg['To'] = recipient

    try:
        with smtplib.SMTP('smtp-mail.outlook.com', 587) as server:
            server.starttls()
            server.login(OUTLOOK_EMAIL, CONTRASEÑA_APLICACION)
            server.sendmail(msg['From'], [msg['To']], msg.as_string())
        flash('Correo electrónico enviado exitosamente!', 'success')
    except Exception as e:
        flash(f'No se pudo enviar el correo electrónico: {str(e)}', 'danger')

# Run the app
if __name__ == '__main__':
    app.run(debug=True)
