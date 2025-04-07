# Transcripteur OneNote WMA vers Word

Ce projet propose un petit programme Python, doté d’une interface graphique (via PyQt5), permettant de :

1. **Sélectionner un fichier audio** – au format `.wma` (par exemple, produit par OneNote) ou tout autre format pris en charge (mp3, wav, m4a, etc.).
2. **Choisir l’emplacement** et le nom du **fichier Word (.docx)** qui contiendra la transcription.
3. **Transcrire localement** le fichier audio grâce à **OpenAI Whisper**, un moteur de reconnaissance vocale performant, même pour des enregistrements de qualité moyenne.
4. **Générer** un document Word (.docx) contenant la transcription.

> **Version du code** : Première version simplifiée (sans différenciation des locuteurs)

---

## Sommaire

- [Fonctionnalités](#fonctionnalités)
- [Installation](#installation)
  - [Python et bibliothèques](#python-et-bibliothèques)
  - [FFmpeg](#ffmpeg)
- [Utilisation](#utilisation)
- [Recommandations](#recommandations)
- [Licence](#licence)
- [Journal des changements](#journal-des-changements)

---

## 1. Fonctionnalités

- Interface graphique simple via PyQt5.
- Conversion automatique de fichiers WMA en WAV (via pydub et FFmpeg).
- Transcription locale avec OpenAI Whisper.
- Utilisation de paramètres de transcription optimisés (`temperature=0.0` et `best_of=3`) pour améliorer la précision et réduire les répétitions.
- Post-traitement pour éliminer les phrases consécutives dupliquées.
- Génération d’un document Word (.docx) contenant le texte transcrit.

---

## 2. Installation

### Python et bibliothèques

Assurez-vous d’avoir **Python 3.8+** installé. Dans un terminal ou une invite de commande, installez les dépendances :

```bash
pip install PyQt5 openai-whisper pydub python-docx
```

- **PyQt5** : Pour l’interface graphique.
- **openai-whisper** : Pour la transcription.
- **pydub** : Pour la conversion audio (WMA -> WAV).
- **python-docx** : Pour générer le document Word.

### FFmpeg

1. Téléchargez [FFmpeg](https://ffmpeg.org/download.html).
2. Décompressez-le dans un dossier, par exemple `C:\ffmpeg`.
3. Ajoutez le dossier `C:\ffmpeg\bin` à la variable d’environnement **PATH**.
4. Vérifiez l’installation en exécutant :

```bash
ffmpeg -version
```

---

## 3. Utilisation

1. Lancez le script avec la commande :

```bash
python transcriber.py
```

2. Dans l’interface graphique qui s’ouvre, choisissez le fichier audio d’entrée et le fichier de sortie (.docx).
3. Cliquez sur **"Transcrire l'audio"**.
4. La transcription est effectuée et le document Word est généré dans l’emplacement spécifié.

---

## 4. Recommandations

- Pour de meilleurs résultats, assurez-vous que l’audio est clair. Si nécessaire, pré-traitez le fichier pour réduire le bruit.
- Les paramètres `temperature=0.0` et `best_of=3` sont utilisés pour limiter les répétitions. Vous pouvez les ajuster selon vos besoins.
- Une relecture manuelle de la transcription est recommandée afin de corriger d’éventuelles erreurs.

---

## 5. Licence

Ce logiciel est distribué sous la **Licence Publique Générale GNU version 3** (GNU GPL v3).

> **Note :** Seule la version anglaise de la GNU GPL v3 a valeur légale officielle. La traduction française est fournie à titre informatif.

---

## 6. Journal des changements

### 2025-04-06 – 18:30 (HE)
- **Version initiale** : Transcription de fichiers audio (WMA, MP3, WAV, M4A) en document Word.
- Implémentation d’une interface graphique simple avec PyQt5.
- Conversion de WMA en WAV via pydub et utilisation de OpenAI Whisper pour la transcription.

### 2025-04-06 – 21:45 (HE)
- **Améliorations de précision** : Passage au modèle "large-v2" de Whisper.
- Ajout des paramètres de transcription (`temperature=0.0` et `best_of=3`) pour réduire les répétitions.
- Implémentation d’un post-traitement pour supprimer les phrases consécutives dupliquées dans la transcription.

