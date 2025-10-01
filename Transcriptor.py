import speech_recognition as sr
from docx import Document
from datetime import datetime
import os

def Transcribir():
    listener = sr.Recognizer()

    try: 
        with sr.Microphone() as source:
            # Obtener la fecha y hora actual en formato día-mes-año
            fecha_y_hora = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
            
            #Escuchar
            print("Escuchando...")
            voice = listener.listen(source)

            # Crear una carpeta llamada "Audios" si no existe
            carpeta_audio = "Audios"
            if not os.path.exists(carpeta_audio):
                os.makedirs(carpeta_audio)  # Crear la carpeta de audios
            
            # Guardar el audio grabado en un archivo .wav dentro de la carpeta "Audios"
            nombre_audio = f"Audio_{fecha_y_hora}.wav"
            ruta_audio = os.path.join(carpeta_audio, nombre_audio)
            with open(ruta_audio, "wb") as audio_file:
                audio_file.write(voice.get_wav_data())  #Guardar los datos en formato WAV
            
            print(f"Audio guardado como: {ruta_audio}")


            #Reconocimiento del audio
            rec = listener.recognize_google(voice, language='es-Es')
            print(rec)
            
            # Crear el nombre del archivo con fecha y hora
            nombre_archivo = f"Transcripcion_{fecha_y_hora}.docx"
            
            # Crear una carpeta llamada "Transcripciones" si no existe
            carpeta_Transcripciones = "Transcripciones"
            if not os.path.exists(carpeta_Transcripciones):
                os.makedirs(carpeta_Transcripciones)  #Crear la carpeta de Transcripciones
            
            # Guardar el archivo en la carpeta
            ruta_completa = os.path.join(carpeta_Transcripciones, nombre_archivo)
            
            # Crear y guardar el archivo .docx
            doc = Document()
            doc.add_paragraph(rec)
            doc.save(ruta_completa)

            print("Archivo guardado como: " + ruta_completa)

    except:
        print("No se logro reconocer el audio, intentelo de nuevo" )