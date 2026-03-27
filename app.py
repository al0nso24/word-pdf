from flask import Flask, render_template, request, send_file
import os
import subprocess

app = Flask(__name__)

#Carpeta donde se guardarán los archivos subidos
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

#Ruta principal
@app.route('/')
def index():
    return render_template('index.html')

#Ruta para manejar la subida de archivos
@app.route('/convert', methods=['POST'])
def convert():
    if 'archivo' not in request.files:
        return "No se envió ningún archivo", 400
    
    archivo = request.files['archivo']

    if archivo.filename == "":
        return "No se ha seleccionado ningún archivo", 400
    

    #Guardar el archivo word
    ruta_word = os.path.join(UPLOAD_FOLDER, archivo.filename)
    archivo.save(ruta_word)

    try:
        #Convertir el archivo word a pdf usando LibreOffice
        subprocess.run(
            [r'C:\Program Files\LibreOffice\program\soffice.exe', 
             '--headless', 
             '--convert-to', 
             'pdf', 
             ruta_word, 
             '--outdir', UPLOAD_FOLDER]
            , check=True)
        
        #Obtener el nombre del archivo pdf generado
        nombre_pdf = os.path.splitext(archivo.filename)[0] + '.pdf'
        ruta_pdf = os.path.join(UPLOAD_FOLDER, nombre_pdf)

        #Verificar que el archivo pdf se haya generado
        if not os.path.exists(ruta_pdf):
            return "Error al convertir el archivo a PDF"
        
        #Enviar el pdf como descarga
        return send_file(ruta_pdf, as_attachment=True)
    
    except subprocess.CalledProcessError:
        return "Error al convertir el archivo"
    

#Ejecutar la aplicación
if __name__ == '__main__':
    app.run(debug=True)