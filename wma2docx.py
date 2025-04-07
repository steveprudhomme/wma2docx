import os
import sys
from PyQt5 import QtWidgets, QtGui  # PyQt5 for the GUI
from pydub import AudioSegment      # Pydub for audio conversion (uses ffmpeg)
import whisper                     # OpenAI Whisper for transcription
from docx import Document          # python-docx for Word document generation

class TranscriptionApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Transcripteur OneNote WMA vers Word")
        self.resize(500, 150)
        
        # Create GUI elements
        self.input_edit = QtWidgets.QLineEdit(self)
        self.input_button = QtWidgets.QPushButton("Choisir un fichier WMA...", self)
        self.output_edit = QtWidgets.QLineEdit(self)
        self.output_button = QtWidgets.QPushButton("Choisir le fichier de sortie...", self)
        self.run_button = QtWidgets.QPushButton("Transcrire l'audio", self)
        self.status_label = QtWidgets.QLabel("Statut : en attente", self)
        
        # Arrange elements in the window using layout
        layout = QtWidgets.QGridLayout(self)
        layout.addWidget(QtWidgets.QLabel("Fichier audio (.wma) :"), 0, 0)
        layout.addWidget(self.input_edit, 0, 1)
        layout.addWidget(self.input_button, 0, 2)
        layout.addWidget(QtWidgets.QLabel("Document Word de sortie :"), 1, 0)
        layout.addWidget(self.output_edit, 1, 1)
        layout.addWidget(self.output_button, 1, 2)
        layout.addWidget(self.run_button, 2, 1)
        layout.addWidget(self.status_label, 3, 0, 1, 3)
        
        # Connect button actions to methods
        self.input_button.clicked.connect(self.select_input_file)
        self.output_button.clicked.connect(self.select_output_file)
        self.run_button.clicked.connect(self.run_transcription)
    
    def select_input_file(self):
        """Open file dialog to select the WMA audio file."""
        options = QtWidgets.QFileDialog.Options()
        # Filter to display only WMA files or all audio files
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Sélectionnez un fichier audio OneNote", "", 
            "Fichiers audio (*.wma *.mp3 *.wav *.m4a);;Tous les fichiers (*)", options=options)
        if file_path:
            self.input_edit.setText(file_path)
    
    def select_output_file(self):
        """Open file dialog to specify the output Word document."""
        options = QtWidgets.QFileDialog.Options()
        # Default extension .docx for output
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Nom du document de transcription", "", 
            "Document Word (*.docx);;Tous les fichiers (*)", options=options)
        if file_path:
            # Ensure the filename has .docx extension
            if not file_path.lower().endswith(".docx"):
                file_path += ".docx"
            self.output_edit.setText(file_path)
    
    def run_transcription(self):
        """Perform the audio transcription and save to Word document."""
        audio_path = self.input_edit.text().strip()
        output_path = self.output_edit.text().strip()
        if not audio_path or not output_path:
            QtWidgets.QMessageBox.warning(self, "Champs manquants",
                                          "Veuillez sélectionner un fichier audio et un fichier de sortie.")
            return
        # Update status
        self.status_label.setText("Statut : transcription en cours...")
        QtGui.QGuiApplication.processEvents()  # refresh UI
        
        try:
            # 1. Conversion du WMA en WAV (si nécessaire)
            temp_wav = None
            if audio_path.lower().endswith(".wma"):
                # Chargement du fichier .wma via pydub (ffmpeg requis)
                audio = AudioSegment.from_file(audio_path, format="asf")
                # Export en WAV PCM 16 bits
                temp_wav = audio_path + "_temp.wav"
                audio.export(temp_wav, format="wav")
                audio_to_transcribe = temp_wav
            else:
                # Si le fichier est déjà dans un format lisible (mp3, wav, etc.)
                audio_to_transcribe = audio_path
            
            # 2. Chargement du modèle Whisper (on peut choisir un modèle plus grand pour plus de précision)
            model = whisper.load_model("medium")  # modèle 'medium' multilingue (équilibre précision/temps)
            
            # 3. Transcription de l'audio en texte (en spécifiant la langue française pour plus de précision)
            result = model.transcribe(audio_to_transcribe, language="fr")
            transcribed_text = result["text"]
            
            # 4. Création du document Word et écriture du texte transcrit
            doc = Document()
            doc.add_paragraph(transcribed_text)
            doc.save(output_path)
            
            # Nettoyage du fichier temporaire si utilisé
            if temp_wav and os.path.exists(temp_wav):
                os.remove(temp_wav)
            
            # Mise à jour du statut et notification de fin
            self.status_label.setText("Statut : Terminé. Transcription sauvegardée.")
            QtWidgets.QMessageBox.information(self, "Transcription terminée",
                                              "Le fichier a été transcrit avec succès !")
        except Exception as e:
            self.status_label.setText("Statut : Échec de la transcription.")
            QtWidgets.QMessageBox.critical(self, "Erreur",
                                           f"Une erreur est survenue : {e}")

# Exécution de l'application
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = TranscriptionApp()
    window.show()
    sys.exit(app.exec_())