import os
import sys
import tempfile
import re
from PyQt5 import QtWidgets, QtGui
from pydub import AudioSegment
import whisper
from docx import Document

def remove_duplicate_sentences(text):
    """
    Supprime les phrases consécutives identiques.
    On découpe le texte sur ". " et on reconstruit en gardant une seule occurrence des phrases consécutives identiques.
    """
    # On découpe le texte en phrases (en utilisant ". " comme séparateur)
    sentences = text.split('. ')
    filtered = []
    for sentence in sentences:
        if not filtered or sentence.strip() != filtered[-1].strip():
            filtered.append(sentence.strip())
    # Reconstruction du texte ; on ajoute un point final s'il manque
    new_text = '. '.join(filtered)
    if new_text and new_text[-1] != '.':
        new_text += '.'
    return new_text

class TranscriptionApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Transcripteur OneNote WMA vers Word")
        self.resize(500, 150)
        
        # Création des widgets
        self.input_edit = QtWidgets.QLineEdit(self)
        self.input_button = QtWidgets.QPushButton("Choisir un fichier audio...", self)
        self.output_edit = QtWidgets.QLineEdit(self)
        self.output_button = QtWidgets.QPushButton("Choisir le fichier de sortie...", self)
        self.run_button = QtWidgets.QPushButton("Transcrire l'audio", self)
        self.status_label = QtWidgets.QLabel("Statut : en attente", self)
        
        # Organisation avec un layout
        layout = QtWidgets.QGridLayout(self)
        layout.addWidget(QtWidgets.QLabel("Fichier audio (.wma, .mp3, .wav, .m4a):"), 0, 0)
        layout.addWidget(self.input_edit, 0, 1)
        layout.addWidget(self.input_button, 0, 2)
        layout.addWidget(QtWidgets.QLabel("Document Word (.docx):"), 1, 0)
        layout.addWidget(self.output_edit, 1, 1)
        layout.addWidget(self.output_button, 1, 2)
        layout.addWidget(self.run_button, 2, 1)
        layout.addWidget(self.status_label, 3, 0, 1, 3)
        
        # Connexions
        self.input_button.clicked.connect(self.select_input_file)
        self.output_button.clicked.connect(self.select_output_file)
        self.run_button.clicked.connect(self.run_transcription)
    
    def select_input_file(self):
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Sélectionnez un fichier audio", "", 
            "Fichiers audio (*.wma *.mp3 *.wav *.m4a);;Tous les fichiers (*)", options=options
        )
        if file_path:
            self.input_edit.setText(file_path)
    
    def select_output_file(self):
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Nom du document de transcription", "", 
            "Document Word (*.docx);;Tous les fichiers (*)", options=options
        )
        if file_path:
            if not file_path.lower().endswith(".docx"):
                file_path += ".docx"
            self.output_edit.setText(file_path)
    
    def run_transcription(self):
        audio_path = self.input_edit.text().strip()
        output_path = self.output_edit.text().strip()
        
        if not audio_path or not output_path:
            QtWidgets.QMessageBox.warning(
                self, "Champs manquants",
                "Veuillez sélectionner un fichier audio et un fichier de sortie."
            )
            return
        
        self.status_label.setText("Statut : Transcription en cours...")
        QtGui.QGuiApplication.processEvents()
        
        try:
            # Conversion en WAV si nécessaire
            temp_wav = None
            ext = os.path.splitext(audio_path)[1].lower()
            if ext == ".wma":
                audio = AudioSegment.from_file(audio_path, format="asf")
                temp_wav = tempfile.mktemp(suffix=".wav")
                audio.export(temp_wav, format="wav")
                final_audio_path = temp_wav
            else:
                final_audio_path = audio_path
            
            # Chargement du modèle Whisper "large-v2" pour une meilleure précision
            model = whisper.load_model("large-v2")
            
            # Transcription avec paramètres ajustés pour réduire les répétitions
            result = model.transcribe(final_audio_path, language="fr", temperature=0.0, best_of=3)
            raw_text = result["text"]
            
            # Post-traitement pour enlever les répétitions consécutives
            processed_text = remove_duplicate_sentences(raw_text)
            
            # Création et sauvegarde du document Word
            doc = Document()
            doc.add_paragraph(processed_text)
            doc.save(output_path)
            
            if temp_wav and os.path.exists(temp_wav):
                os.remove(temp_wav)
            
            self.status_label.setText("Statut : Transcription terminée.")
            QtWidgets.QMessageBox.information(
                self, "Transcription terminée",
                "La transcription a été réalisée avec succès."
            )
        except Exception as e:
            self.status_label.setText("Statut : Échec de la transcription.")
            QtWidgets.QMessageBox.critical(self, "Erreur", str(e))

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = TranscriptionApp()
    window.show()
    sys.exit(app.exec_())
