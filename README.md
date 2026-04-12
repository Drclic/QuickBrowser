# PDF Quick Browser v3

Navigateur de fichiers avec preview PDF dans une fenêtre séparée, inspiré de Path Finder (macOS).

## Architecture

L'application se compose de **deux fenêtres** positionnées côte à côte :

- **Fenêtre Explorateur** (gauche) : arborescence de dossiers + liste de fichiers complète, comme l'Explorateur Windows
- **Fenêtre Preview** (droite) : affichage natif du PDF sélectionné via QPdfView

Quand vous naviguez dans l'explorateur et sélectionnez un PDF, il s'affiche instantanément dans la fenêtre de preview.

## Fonctionnalités

### Explorateur
- Arborescence de dossiers complète (panneau gauche)
- Liste de tous les fichiers du dossier courant (panneau droit)
- Double-clic pour entrer dans un dossier
- Navigation avant/arrière avec historique
- Barre de chemin cliquable et éditable (Ctrl+L)
- Tri par colonnes (nom, taille, date, type)
- Clic droit : ouvrir dans l'Explorateur, copier chemin/nom

### Preview PDF
- Rendu natif ultra-rapide (QPdfView)
- Cache de 10 documents en mémoire
- Préchargement des PDF adjacents
- Navigation par pages (▲/▼)
- Zoom : ajuster à la largeur, à la page, ou zoom libre

## Raccourcis clavier

### Explorateur
| Raccourci       | Action                        |
|-----------------|-------------------------------|
| `↑` / `↓`      | Naviguer dans la liste        |
| `Entrée`        | Entrer dans le dossier        |
| `Retour arrière`| Dossier parent                |
| `Alt+←`         | Précédent (historique)        |
| `Alt+→`         | Suivant (historique)          |
| `Alt+↑`         | Dossier parent                |
| `Ctrl+L`        | Éditer le chemin              |
| `Ctrl+E`        | Ouvrir dans l'Explorateur     |
| `F5`            | Rafraîchir                    |
| `Page↑/↓`       | Sauter de 10 fichiers         |

### Preview
| Raccourci       | Action                        |
|-----------------|-------------------------------|
| `+` / `−`       | Zoom avant / arrière          |
| `Ctrl+0`        | Ajuster à la largeur          |

## Build

### GitHub Actions
Pusher sur GitHub → Actions → Run workflow → Télécharger l'artefact

### Local
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python pdf_quick_browser.py          # tester
pyinstaller pdf_quick_browser.spec   # builder
```
