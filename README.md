# Meeting SumUp

Automatisation Windows qui **surveille un dossier** de transcriptions de réunion et **génère automatiquement** une synthèse structurée au format **Word (.docx)**, avec une **notification Windows cliquable** à la fin (ouvre le fichier généré).

## Intérêt

- **0 friction**: tu déposes un fichier → tu récupères un `.docx`
- **Synthèse actionnable**: résumé, décisions, actions, points clés
- **Industrialisation**: tourne en fond avec un watcher

## Fonctionnement

- `watcher.py` surveille `WATCH_FOLDER` pour les extensions `.txt`, `.vtt`, `.docx`
- À chaque nouveau fichier:
  - `process.py` lit la transcription
  - appelle un LLM (Claude ou Gemini selon la config)
  - génère un `.docx` dans `OUTPUT_FOLDER`
  - envoie une notification Windows (clic → ouvre le fichier)

## Prérequis

- Windows 10/11
- Python 3.10+ recommandé

## Installation

```bash
git clone <ton-repo>
cd meeting-sumup
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

## Configuration

1) Créer ton fichier `.env` à partir de l’exemple:

```bash
copy .env.example .env
```

2) Renseigner au minimum:

- `WATCH_FOLDER` (dossier surveillé)
- `OUTPUT_FOLDER` (dossier de sortie)
- une API active:
  - **Claude**: `USE_CLAUDE=true` + `CLAUDE_API_KEY=...`
  - ou **Gemini**: `USE_GEMINI=true` + `GEMINI_API_KEY=...`

## Lancer

```bash
py watcher.py
```

Dépose ensuite une transcription dans `WATCH_FOLDER`.

