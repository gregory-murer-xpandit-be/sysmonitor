# SysMonitor

Dashboard système Windows 11 en temps réel — réseau, applications installées et patches/mises à jour, dans une interface desktop sombre et moderne.

## Fonctionnalités

**Vue d'ensemble**
- CPU, RAM, disque C: en temps réel
- Mini-graphes bande passante réseau et CPU/RAM (60 secondes glissantes)
- Timeline des patches des 12 derniers mois
- Top 5 applications par consommation mémoire

**Onglet Réseau**
- Connexions actives en temps réel (toutes les secondes)
- Répartition HTTP / HTTPS / autres protocoles
- Bande passante montante/descendante en KB/s
- Tableau filtrable des connexions établies (IP, port, processus)

**Onglet Applications**
- Lecture directe du registre Windows (< 1 s, sans WMI)
- 4 clés lues en parallèle : HKLM x64/x32 + HKCU x64/x32
- Noms en orange = application active en mémoire
- Mémoire en rouge = > 500 MB
- Tri par nom, mémoire active, taille disque, version
- Recherche en temps réel + export CSV

**Onglet Patches / MAJ**
- État Windows Update (à jour / en attente)
- Liste des hotfixes installés (Win32_QuickFixEngineering)
- Timeline graphique des 12 derniers mois
- Lien direct vers les paramètres Windows Update

## Prérequis

- Windows 11
- Python 3.11+
- Droits administrateur (pour la lecture des connexions réseau)

## Installation

```bash
git clone https://github.com/gregory-murer-xpandit-be/sysmonitor.git
cd sysmonitor
pip install -r requirements.txt
```

## Lancement

```bash
python sysmonitor.py
```

> Lancer en tant qu'administrateur pour accès complet aux connexions réseau.

## Stack technique

| Librairie | Usage |
|-----------|-------|
| PyQt6 | Interface graphique |
| pyqtgraph | Graphes temps réel |
| psutil | CPU, RAM, réseau, processus |
| pywin32 | API Windows |
| winreg | Lecture registre (stdlib) |
| wmi | Patches et infos OS |

## Versions

| Version | Description |
|---------|-------------|
| v0.1 | Dashboard réseau / applications / patches — Windows 11 |

## Auteur

Gregory Murer
