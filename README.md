# Transcripteur OneNote WMA vers Word

Ce projet propose un petit programme Python, doté d’une interface graphique (via PyQt5), permettant de :

1. **Sélectionner un fichier audio** – au format `.wma` (par exemple, produit par OneNote) ou tout autre format pris en charge (mp3, wav, m4a, etc.).  
2. **Choisir l’emplacement** et le nom du **fichier Word (.docx)** qui contiendra la transcription.  
3. **Transcrire localement** le fichier audio grâce à **OpenAI Whisper**, un moteur de reconnaissance vocale hautement performant, même pour des enregistrements de qualité moyenne.  
4. **Générer** un document Word `.docx` contenant le texte transcrit.

> **Version du code** : Première version (sans barre de progression ni signal sonore)

---

## Sommaire

- [1. Fonctionnalités](#1-fonctionnalités)  
- [2. Installation](#2-installation)  
  - [2.1. Python et bibliothèques](#21-python-et-bibliothèques)  
  - [2.2. FFmpeg](#22-ffmpeg)  
- [3. Utilisation](#3-utilisation)  
- [4. Recommandations](#4-recommandations)  
- [5. Licence](#5-licence)

---

## 1. Fonctionnalités

- **Interface graphique simple** via PyQt5.  
- **Conversion automatique** WMA -> WAV si nécessaire (via Pydub + FFmpeg).  
- **Transcription locale** hors-ligne (OpenAI Whisper) – aucune donnée envoyée en ligne.  
- **Export au format .docx** : insertion du texte transcrit dans un document Word.  
- **Support multilingue** (on force par défaut `language="fr"` dans le code).

---

## 2. Installation

### 2.1. Python et bibliothèques

1. **Python 3.8+** : Vérifiez que Python est installé (par exemple [https://www.python.org/downloads/](https://www.python.org/downloads/)).
2. Dans un terminal ou une invite de commande, installez les dépendances avec :

    pip install PyQt5 openai-whisper pydub python-docx

   - **PyQt5** : pour l’interface graphique  
   - **openai-whisper** : pour la reconnaissance vocale (nécessite PyTorch)  
   - **pydub** : pour la conversion audio WMA -> WAV  
   - **python-docx** : pour générer le fichier Word

3. Assurez-vous que **PyTorch** est correctement installé (souvent récupéré via l’installation de `openai-whisper`).  
   - Si vous disposez d’un GPU Nvidia, vous pouvez installer PyTorch avec le support CUDA pour accélérer la transcription.  

### 2.2. FFmpeg

1. Téléchargez [FFmpeg](https://ffmpeg.org/download.html).  
2. Décompressez-le dans un dossier, par exemple `C:\ffmpeg`.  
3. Ajoutez le dossier `C:\ffmpeg\bin` à la variable d’environnement **PATH** de Windows.  
4. Vérifiez l’installation :

    ffmpeg -version

---

## 3. Utilisation

1. **Lancer le script** :

    python convert_wma_txt.py

   Une fenêtre PyQt5 apparaît.

2. **Choix du fichier d’entrée** :  
   - Cliquez sur “Choisir un fichier WMA...”.  
   - Sélectionnez un fichier `.wma`, `.mp3`, `.wav` ou `.m4a`.  

3. **Choix du fichier de sortie** :  
   - Cliquez sur “Choisir le fichier de sortie...”.  
   - Entrez un nom de fichier (ex. `transcription.docx`) et validez.  

4. **Exécuter la transcription** :  
   - Cliquez sur “Transcrire l’audio”.  
   - Le programme convertit éventuellement le fichier en WAV, charge Whisper en français, puis crée le `.docx`.  
   - Un message de confirmation “Transcription terminée” apparaît en cas de succès.

---

## 4. Recommandations

- **Qualité audio** : Si l’audio est très bruité, vous pouvez d’abord le nettoyer (par exemple avec Audacity) pour améliorer le résultat de la transcription.  
- **Modèle Whisper** : Le script utilise `medium` pour un équilibre rapidité-précision. Vous pouvez modifier `model = whisper.load_model("medium")` pour `tiny`, `small` ou `large`.  
- **Langue** : Nous utilisons `language="fr"` pour l’audio en français. Si nécessaire, modifiez ou supprimez ce paramètre pour l’adapter à d’autres langues.  
- **Temps de traitement** : Sur un CPU, la transcription peut être lente pour des audios longs. L’utilisation d’un GPU accélère grandement le processus (nécessite l’installation de PyTorch avec CUDA).  
- **Relecture** : Même si Whisper est performant, une relecture manuelle reste conseillée pour valider la transcription (noms propres, abréviations, termes spécifiques…).

---

## 5. Licence

Ce logiciel est distribué sous la **Licence Publique Générale GNU version 3** (GNU GPL v3).  

> **Note importante :** Seule la version anglaise de la GNU GPL v3 est considérée comme ayant valeur légale officielle. La traduction française ci-dessous est fournie à titre informatif.  

### GNU GENERAL PUBLIC LICENSE  
#### Version 3, 29 juin 2007 (Traduction française - non officielle)

<details>
<summary>Afficher la traduction française (non officielle)</summary>

*(Insérez ici le texte complet de la traduction française de la GPL v3, ou un résumé, tout en renvoyant à la version anglaise officielle.)*

**Version anglaise officielle** : [GNU GPL v3 sur gnu.org](https://www.gnu.org/licenses/gpl-3.0.en.html)

</details>

---

**Auteur** : Steve Prud’homme  
**Dernière mise à jour** : 2025-04-06