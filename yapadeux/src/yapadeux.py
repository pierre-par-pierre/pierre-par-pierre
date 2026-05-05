"""
yapadeux — y'a pas deux

Détecte les fichiers doublons dans tes dossiers Downloads (récursif) et
Desktop (racine seule), les déplace dans un dossier "_doublons-detectes/"
sur le Bureau pour que tu puisses vérifier et supprimer toi-même.

Lancement prévu : double-clic sur LANCER.vbs (qui appelle ce script via
pythonw.exe pour que rien ne s'affiche en console).

Stack : Python 3.10+ stdlib uniquement (Windows).

Note UX : si l'Explorateur Windows n'affiche pas immédiatement les
nouveaux sous-dossiers numérotés au fil de leur création (cache de vue
Windows), appuie sur F5 dans la fenêtre pour rafraîchir.

Mécanismes de contrôle :
- STOP.vbs crée une sentinelle sur le Desktop, lue avant chaque groupe.
- Un verrou (lock file) empêche deux instances de tourner en parallèle.
- RESTAURER.vbs (→ src/restaurer.py) annule un tri en lisant la section
  MAPPING écrite en bas du _RAPPORT.txt.
"""

import ctypes
import hashlib
import os
import shutil
import sys
import time
import traceback
from collections import defaultdict
from datetime import datetime
from pathlib import Path


# ─────────────────────────────────────────────────────────────────────
# Constantes (à ajuster pour les tests, voir bloc plus bas)
# ─────────────────────────────────────────────────────────────────────

MODE_TEST = False
CHEMIN_DOWNLOADS_TEST: Path | None = None
CHEMIN_DESKTOP_TEST: Path | None = None

DELAI_PROGRESSIF = 0.4          # secondes entre chaque groupe affiché
DELAI_PROGRESSIF_RAPIDE = 0.15  # accéléré au-delà du seuil
SEUIL_BASCULE_RAPIDE = 30       # nombre de groupes avant accélération

TAILLE_CHUNK_MD5 = 64 * 1024    # 64 Ko par chunk pour le hash
INTERVALLE_MAJ_RAPPORT = 50     # rafraîchit le rapport tous les N hash

# Garde-fou : un dossier source n'est accepté que s'il contient l'un
# de ces 4 mots-clés. Empêche un MODE_TEST mal configuré de scanner
# n'importe où.
MOTS_CLES_AUTORISES = ("Downloads", "Téléchargements", "Desktop", "Bureau")

# Préfixes de dossiers à exclure du scan récursif Downloads (pour ne
# pas s'auto-scanner si l'utilisateur a copié un ancien tri).
PREFIXES_EXCLUS = ("_doublons-detectes", "_AUCUN_DOUBLON", "_ERREUR_YAPADEUX")

# Sentinelle d'interruption (créée par STOP.vbs sur le Desktop)
def _desktop_pour_signaux() -> Path:
    """Desktop utilisé pour la sentinelle/le verrou. En mode test on
    utilise le faux Desktop pour ne pas polluer le vrai."""
    if MODE_TEST and CHEMIN_DESKTOP_TEST is not None:
        return CHEMIN_DESKTOP_TEST
    return Path.home() / "Desktop"


NOM_SENTINELLE_STOP = "_yapadeux-stop.signal"
NOM_VERROU_RUN = "_yapadeux-running.lock"

# Marqueurs du mapping technique en bas du _RAPPORT.txt (le séparateur
# `|` est interdit dans les noms de fichiers Windows, donc safe).
MARQUEUR_MAPPING_DEBUT = "MAPPING:"
MARQUEUR_MAPPING_FIN = "END_MAPPING"
SEPARATEUR_MAPPING = "|"

# Stocke en mémoire la dernière popup (pour les tests). Inhibe la vraie
# fenêtre Windows si MODE_TEST=True.
derniere_popup: tuple[str, str] | None = None


# ─────────────────────────────────────────────────────────────────────
# Helpers Windows (popup, vérification de processus, verrou)
# ─────────────────────────────────────────────────────────────────────

# Constantes Win32
_MB_OK = 0x0
_MB_ICONWARNING = 0x30
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_STILL_ACTIVE = 259  # GetExitCodeProcess renvoie 259 quand le process tourne


def afficher_popup(titre: str, message: str):
    """Affiche une fenêtre Windows MessageBoxW. En mode test, mémorise
    juste l'appel dans `derniere_popup` (pas de fenêtre)."""
    global derniere_popup
    derniere_popup = (titre, message)
    if MODE_TEST:
        return
    try:
        ctypes.windll.user32.MessageBoxW(0, message, titre, _MB_OK | _MB_ICONWARNING)
    except OSError:
        # Si on n'est pas sur Windows (cas anormal pour ce projet),
        # on log dans un fichier et on continue.
        pass


def processus_actif(pid: int) -> bool:
    """True si un processus avec ce PID tourne encore.
    Utilise OpenProcess + GetExitCodeProcess via ctypes."""
    if pid <= 0:
        return False
    try:
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not handle:
            return False
        try:
            code = ctypes.c_ulong()
            ok = kernel32.GetExitCodeProcess(handle, ctypes.byref(code))
            if not ok:
                return False
            return code.value == _STILL_ACTIVE
        finally:
            kernel32.CloseHandle(handle)
    except OSError:
        return False


def acquerir_verrou(verrou: Path) -> bool:
    """
    Tente d'acquérir le verrou. Retourne True si OK, False si une
    instance tourne déjà (popup affichée).
    Si le verrou est orphelin (PID mort), il est écrasé.
    """
    if verrou.exists():
        try:
            pid_existant = int(verrou.read_text(encoding="utf-8").strip())
        except (ValueError, OSError):
            pid_existant = 0
        if pid_existant and processus_actif(pid_existant):
            afficher_popup(
                "yapadeux",
                "Un tri yapadeux est déjà en cours.\n\n"
                "Attends qu'il se termine, ou double-clique sur "
                "STOP.vbs pour l'interrompre.",
            )
            return False
        # Verrou orphelin : on l'écrase
        try:
            verrou.unlink()
        except OSError:
            pass
    try:
        verrou.write_text(str(os.getpid()), encoding="utf-8")
    except OSError:
        return False
    return True


def liberer_verrou(verrou: Path):
    """Supprime le verrou s'il existe. Idempotent."""
    try:
        if verrou.exists():
            verrou.unlink()
    except OSError:
        pass


# ─────────────────────────────────────────────────────────────────────
# Petite classe pour tagger chaque fichier avec sa source
# ─────────────────────────────────────────────────────────────────────

class Fichier:
    """Représente un fichier scanné, avec sa source (DOWNLOADS / DESKTOP)."""

    __slots__ = ("chemin", "tag", "taille", "mtime", "md5")

    def __init__(self, chemin: Path, tag: str):
        self.chemin = chemin
        self.tag = tag
        st = chemin.stat()
        self.taille = st.st_size
        self.mtime = st.st_mtime
        self.md5: str | None = None


# ─────────────────────────────────────────────────────────────────────
# Détection des dossiers source
# ─────────────────────────────────────────────────────────────────────

def detecter_desktop() -> Path | None:
    """Retourne le Desktop de l'utilisateur courant, ou None."""
    home = Path.home()
    for nom in ("Desktop", "Bureau"):
        candidat = home / nom
        if candidat.is_dir():
            return candidat
    return None


def detecter_downloads() -> Path | None:
    """Retourne le Downloads de l'utilisateur courant, ou None."""
    home = Path.home()
    for nom in ("Downloads", "Téléchargements"):
        candidat = home / nom
        if candidat.is_dir():
            return candidat
    return None


def est_chemin_autorise(p: Path) -> bool:
    """Vérifie qu'un chemin contient au moins un mot-clé autorisé."""
    s = str(p)
    return any(mot in s for mot in MOTS_CLES_AUTORISES)


# ─────────────────────────────────────────────────────────────────────
# Walk + hashage
# ─────────────────────────────────────────────────────────────────────

def doit_etre_exclu(nom: str) -> bool:
    """True si un nom de dossier commence par un des préfixes interdits."""
    return any(nom.startswith(p) for p in PREFIXES_EXCLUS)


def scanner(racine: Path, tag: str, recursif: bool, erreurs: list[str]) -> list[Fichier]:
    """
    Liste les fichiers de `racine` taggés avec `tag`.
    - recursif=True  → descend dans tous les sous-dossiers (sauf préfixes exclus).
    - recursif=False → racine uniquement, ne descend pas.
    """
    fichiers: list[Fichier] = []
    if recursif:
        for dossier, sous_dossiers, noms_fichiers in os.walk(racine):
            # Filtrage in-place pour ne pas descendre dans les sorties d'anciens runs
            sous_dossiers[:] = [d for d in sous_dossiers if not doit_etre_exclu(d)]
            for nom in noms_fichiers:
                chemin = Path(dossier) / nom
                try:
                    if chemin.is_file():
                        fichiers.append(Fichier(chemin, tag))
                except OSError as e:
                    erreurs.append(f"Lecture impossible : {chemin} ({e})")
    else:
        # Racine uniquement
        try:
            entries = list(racine.iterdir())
        except OSError as e:
            erreurs.append(f"Lecture impossible : {racine} ({e})")
            return fichiers
        for chemin in entries:
            try:
                if chemin.is_file():
                    fichiers.append(Fichier(chemin, tag))
            except OSError as e:
                erreurs.append(f"Lecture impossible : {chemin} ({e})")
    return fichiers


def hash_md5(p: Path) -> str | None:
    """MD5 par chunks de 64 Ko. None si lecture impossible."""
    h = hashlib.md5()
    try:
        with open(p, "rb") as f:
            while True:
                bloc = f.read(TAILLE_CHUNK_MD5)
                if not bloc:
                    break
                h.update(bloc)
        return h.hexdigest()
    except OSError:
        return None


# ─────────────────────────────────────────────────────────────────────
# Arbitrage des doublons
# ─────────────────────────────────────────────────────────────────────

def grouper_par_taille(fichiers: list[Fichier]) -> list[list[Fichier]]:
    """Groupes de >= 2 fichiers de même taille."""
    par_taille: dict[int, list[Fichier]] = defaultdict(list)
    for f in fichiers:
        par_taille[f.taille].append(f)
    return [g for g in par_taille.values() if len(g) >= 2]


def grouper_par_hash(fichiers: list[Fichier]) -> list[list[Fichier]]:
    """Groupes de >= 2 fichiers de même MD5 (ignore les fichiers sans hash)."""
    par_hash: dict[str, list[Fichier]] = defaultdict(list)
    for f in fichiers:
        if f.md5 is not None:
            par_hash[f.md5].append(f)
    return [g for g in par_hash.values() if len(g) >= 2]


def elire_gagnant(groupe: list[Fichier]) -> Fichier:
    """
    Le gagnant ne bouge jamais. Règle :
    - S'il y a au moins un fichier DESKTOP, le gagnant = DESKTOP plus ancien.
    - Sinon, gagnant = plus ancien parmi les DOWNLOADS.
    """
    desktops = [f for f in groupe if f.tag == "DESKTOP"]
    if desktops:
        return min(desktops, key=lambda f: f.mtime)
    return min(groupe, key=lambda f: f.mtime)


# ─────────────────────────────────────────────────────────────────────
# Déplacement (avec gestion des collisions de noms)
# ─────────────────────────────────────────────────────────────────────

def nom_destination_unique(dossier: Path, nom_propose: str) -> str:
    """Si dossier/nom_propose existe déjà, suffixe " (1)", " (2)"…"""
    cible = dossier / nom_propose
    if not cible.exists():
        return nom_propose
    base = Path(nom_propose).stem
    ext = Path(nom_propose).suffix
    n = 1
    while True:
        candidat = f"{base} ({n}){ext}"
        if not (dossier / candidat).exists():
            return candidat
        n += 1


# ─────────────────────────────────────────────────────────────────────
# Génération du _RAPPORT.txt
# ─────────────────────────────────────────────────────────────────────

LIGNE = "─" * 39
LIGNE_DOUBLE = "═" * 39


def ecrire_rapport_initial(chemin_rapport: Path):
    """Premier rapport, écrit AVANT le scan."""
    contenu = (
        f"{LIGNE_DOUBLE}\n"
        "   YAPADEUX — y'a pas deux\n"
        f"{LIGNE_DOUBLE}\n"
        "Scan en cours...\n\n"
        "Ouvre ce fichier à la fin pour voir le bilan complet.\n"
        "Les sous-dossiers numérotés vont apparaître au fur et à mesure.\n"
        "(Si l'Explorateur ne se rafraîchit pas tout seul, appuie sur F5.)\n"
    )
    chemin_rapport.write_text(contenu, encoding="utf-8")


def ecrire_rapport_progression(chemin_rapport: Path, hashes_faits: int, total: int):
    """Rapport pendant le hashage."""
    contenu = (
        f"{LIGNE_DOUBLE}\n"
        "   YAPADEUX — y'a pas deux\n"
        f"{LIGNE_DOUBLE}\n"
        f"Hashage en cours : {hashes_faits}/{total} fichiers...\n\n"
        "Ouvre ce fichier à la fin pour voir le bilan complet.\n"
        "(Tu peux rafraîchir cette fenêtre avec F5.)\n"
    )
    chemin_rapport.write_text(contenu, encoding="utf-8")


def formater_taille(octets: int) -> str:
    """Formate une taille en Ko/Mo/Go pour le bilan."""
    for unite, seuil in (("Go", 1024**3), ("Mo", 1024**2), ("Ko", 1024)):
        if octets >= seuil:
            return f"{octets / seuil:.1f} {unite}"
    return f"{octets} octets"


def ecrire_rapport_final(
    chemin_rapport: Path,
    chemin_dl: Path | None,
    chemin_desk: Path | None,
    fichiers_scannes: int,
    groupes_traites: list[tuple[int, list[Fichier], Fichier, list[tuple[Fichier, str]]]],
    erreurs: list[str],
    fichiers_deplaces_total: int,
    espace_concerne: int,
    deplacements_effectues: list[tuple[Path, Path]],
    interrompu: bool,
    debut_scan: datetime,
    nb_groupes_prevus: int,
):
    """
    Écrit le rapport final complet, avec section MAPPING en bas.
    Si interrompu=True, ajoute le bandeau d'interruption.
    """
    lignes: list[str] = []
    lignes.append(LIGNE_DOUBLE)
    lignes.append("   YAPADEUX — y'a pas deux")
    if interrompu:
        lignes.append("   Bilan du scan (INTERROMPU)")
    else:
        lignes.append("   Bilan du scan")
    lignes.append(LIGNE_DOUBLE)
    lignes.append("")
    if interrompu:
        lignes.append("⚠ SCAN INTERROMPU PAR L'UTILISATEUR")
        lignes.append("")
    lignes.append(f"Date du scan       : {debut_scan:%Y-%m-%d %H:%M}")
    if interrompu:
        lignes.append(f"Interrompu à       : {datetime.now():%Y-%m-%d %H:%M}")
        lignes.append(
            f"Groupes traités    : {len(groupes_traites)} / "
            f"{nb_groupes_prevus} prévus"
        )
    lignes.append("Dossiers scannés   :")
    if chemin_dl is not None:
        lignes.append(f"  ✓ Downloads : {chemin_dl} (récursif)")
    else:
        lignes.append("  ✗ Downloads : non trouvé")
    if chemin_desk is not None:
        lignes.append(f"  ✓ Desktop   : {chemin_desk} (racine uniquement)")
    else:
        lignes.append("  ✗ Desktop   : non trouvé")
    lignes.append("")
    lignes.append(LIGNE)
    lignes.append("Bilan")
    lignes.append(LIGNE)
    lignes.append(f"Fichiers scannés     : {fichiers_scannes}")
    lignes.append(f"Groupes de doublons  : {len(groupes_traites)}")
    lignes.append(f"Fichiers déplacés    : {fichiers_deplaces_total}")
    lignes.append(f"Espace concerné      : {formater_taille(espace_concerne)}")
    lignes.append("")
    if interrompu:
        lignes.append(
            "Pour annuler complètement et remettre tes fichiers en place,"
        )
        lignes.append("double-clique sur RESTAURER.vbs.")
        lignes.append("")
    lignes.append(LIGNE)
    lignes.append("Détail par groupe")
    lignes.append(LIGNE)
    lignes.append("")
    if not groupes_traites:
        lignes.append("Aucun doublon trouvé.")
        lignes.append("")
    else:
        for num, _groupe, gagnant, perdants in groupes_traites:
            nom_aff = gagnant.chemin.name
            md5_court = (gagnant.md5 or "?")[:12] + "..."
            mtime_str = datetime.fromtimestamp(gagnant.mtime).strftime("%Y-%m-%d")
            lignes.append(f'Groupe {num:02d} — "{nom_aff}"')
            lignes.append(f"  MD5     : {md5_court}")
            lignes.append(f"  Gagnant : {gagnant.tag} — {gagnant.chemin}")
            lignes.append(f"            (modifié {mtime_str})")
            lignes.append("  Déplacés :")
            for perdant, nom_final in perdants:
                lignes.append(f"    - {perdant.tag} — {perdant.chemin}")
                lignes.append(f"      → {nom_final}")
            lignes.append("")
    lignes.append(LIGNE)
    lignes.append("Erreurs rencontrées")
    lignes.append(LIGNE)
    if erreurs:
        for e in erreurs:
            lignes.append(f"- {e}")
    else:
        lignes.append("Aucune erreur.")
    lignes.append("")
    lignes.append(LIGNE)
    lignes.append("Que faire maintenant ?")
    lignes.append(LIGNE)
    lignes.append("1. Ouvre chaque sous-dossier numéroté.")
    lignes.append("2. Vérifie que les fichiers sont bien des doublons que tu veux supprimer.")
    lignes.append("3. Si OK, supprime le sous-dossier (Maj+Suppr pour bypass corbeille).")
    lignes.append("4. Le fichier \"gagnant\" reste à sa place d'origine, intact.")
    lignes.append("")
    # Section technique pour restaurer.py
    lignes.append(LIGNE)
    lignes.append("# SECTION TECHNIQUE — NE PAS MODIFIER")
    lignes.append("# Utilisée par restaurer.py pour remettre les fichiers en place.")
    lignes.append(LIGNE)
    lignes.append(MARQUEUR_MAPPING_DEBUT)
    for origine, destination in deplacements_effectues:
        lignes.append(f"{origine}{SEPARATEUR_MAPPING}{destination}")
    lignes.append(MARQUEUR_MAPPING_FIN)
    lignes.append("")

    chemin_rapport.write_text("\n".join(lignes), encoding="utf-8")


# ─────────────────────────────────────────────────────────────────────
# Coeur de l'algo
# ─────────────────────────────────────────────────────────────────────

def construire_nom_dossier_sortie(desktop: Path) -> Path:
    """Suffixe avec timestamp si _doublons-detectes/ existe déjà."""
    base = desktop / "_doublons-detectes"
    if not base.exists():
        return base
    suffixe = datetime.now().strftime("%Y-%m-%d_%Hh%M")
    return desktop / f"_doublons-detectes_{suffixe}"


def ouvrir_explorateur(chemin: Path):
    """Ouvre l'Explorateur Windows. Pas en mode test."""
    if MODE_TEST:
        return
    try:
        os.startfile(str(chemin))  # type: ignore[attr-defined]
    except OSError:
        # Antivirus ou OS non-Windows : tant pis, le rapport est écrit.
        pass


def main():
    debut_scan = datetime.now()
    erreurs: list[str] = []

    # ───── Phase 0 : détection + garde-fou + verrou

    if MODE_TEST:
        chemin_dl = CHEMIN_DOWNLOADS_TEST
        chemin_desk = CHEMIN_DESKTOP_TEST
    else:
        chemin_dl = detecter_downloads()
        chemin_desk = detecter_desktop()

    if chemin_dl is not None and not est_chemin_autorise(chemin_dl):
        erreurs.append(f"Chemin Downloads refusé (mot-clé absent) : {chemin_dl}")
        chemin_dl = None
    if chemin_desk is not None and not est_chemin_autorise(chemin_desk):
        erreurs.append(f"Chemin Desktop refusé (mot-clé absent) : {chemin_desk}")
        chemin_desk = None

    if chemin_dl is None and chemin_desk is None:
        message = (
            "yapadeux n'a pas trouvé tes dossiers Downloads et Desktop.\n"
            "Vérifie qu'ils existent dans ton dossier utilisateur :\n"
            f"  {Path.home()}\n"
        )
        if erreurs:
            message += "\nDétails :\n" + "\n".join(f"  - {e}" for e in erreurs)
        cible_erreur = (chemin_desk or Path.home() / "Desktop")
        cible_erreur.mkdir(parents=True, exist_ok=True)
        (cible_erreur / "_ERREUR_YAPADEUX.txt").write_text(message, encoding="utf-8")
        ouvrir_explorateur(cible_erreur)
        sys.exit(1)

    # On a besoin d'un Desktop pour poser les signaux et la sortie.
    desktop_pour_sortie = chemin_desk if chemin_desk is not None else (Path.home() / "Desktop")
    desktop_pour_sortie.mkdir(parents=True, exist_ok=True)

    # Verrou anti-double-lancement
    verrou = desktop_pour_sortie / NOM_VERROU_RUN
    sentinelle_stop = desktop_pour_sortie / NOM_SENTINELLE_STOP

    if not acquerir_verrou(verrou):
        sys.exit(1)

    # Nettoyer une éventuelle sentinelle résiduelle d'un run précédent
    try:
        if sentinelle_stop.exists():
            sentinelle_stop.unlink()
    except OSError:
        pass

    try:
        _executer_tri(
            chemin_dl=chemin_dl,
            chemin_desk=chemin_desk,
            desktop_pour_sortie=desktop_pour_sortie,
            sentinelle_stop=sentinelle_stop,
            erreurs=erreurs,
            debut_scan=debut_scan,
        )
    finally:
        # Toujours libérer le verrou et nettoyer une sentinelle laissée
        # (par exemple si l'utilisateur clique STOP pile à la dernière seconde).
        liberer_verrou(verrou)
        try:
            if sentinelle_stop.exists():
                sentinelle_stop.unlink()
        except OSError:
            pass


def _executer_tri(
    chemin_dl: Path | None,
    chemin_desk: Path | None,
    desktop_pour_sortie: Path,
    sentinelle_stop: Path,
    erreurs: list[str],
    debut_scan: datetime,
):
    """Exécute le tri proprement dit (extrait de main pour le finally)."""

    # ───── Phase 1 : préparation visible immédiatement

    dossier_sortie = construire_nom_dossier_sortie(desktop_pour_sortie)
    dossier_sortie.mkdir(parents=True, exist_ok=True)
    chemin_rapport = dossier_sortie / "_RAPPORT.txt"
    ecrire_rapport_initial(chemin_rapport)
    ouvrir_explorateur(dossier_sortie)

    # ───── Phase 2 : scan + arbitrage

    fichiers: list[Fichier] = []
    if chemin_dl is not None:
        fichiers.extend(scanner(chemin_dl, "DOWNLOADS", True, erreurs))
    if chemin_desk is not None:
        # Desktop : racine seule, on ne descend pas
        fichiers.extend(scanner(chemin_desk, "DESKTOP", False, erreurs))

    fichiers_scannes_total = len(fichiers)

    groupes_taille = grouper_par_taille(fichiers)
    candidats_hash = [f for groupe in groupes_taille for f in groupe]
    total_a_hasher = len(candidats_hash)

    for i, f in enumerate(candidats_hash, start=1):
        h = hash_md5(f.chemin)
        if h is None:
            erreurs.append(f"Hash impossible : {f.chemin}")
        else:
            f.md5 = h
        if i % INTERVALLE_MAJ_RAPPORT == 0:
            ecrire_rapport_progression(chemin_rapport, i, total_a_hasher)

    groupes_hash = grouper_par_hash(candidats_hash)
    groupes_hash.sort(key=lambda g: g[0].taille, reverse=True)
    nb_groupes_prevus = len(groupes_hash)

    # ───── Phase 3 : déplacement progressif (avec test STOP)

    groupes_traites: list[tuple[int, list[Fichier], Fichier, list[tuple[Fichier, str]]]] = []
    deplacements_effectues: list[tuple[Path, Path]] = []
    fichiers_deplaces_total = 0
    espace_concerne = 0
    interrompu = False

    for index, groupe in enumerate(groupes_hash, start=1):
        # Checkpoint STOP avant chaque groupe
        if sentinelle_stop.exists():
            try:
                sentinelle_stop.unlink()
            except OSError:
                pass
            interrompu = True
            break

        gagnant = elire_gagnant(groupe)
        perdants = [f for f in groupe if f is not gagnant]

        nom_dossier_groupe = f"{index:02d}_{gagnant.chemin.name}"
        dossier_groupe = dossier_sortie / nom_dossier_groupe
        try:
            dossier_groupe.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            erreurs.append(f"Création dossier groupe impossible : {dossier_groupe} ({e})")
            continue

        perdants_deplaces: list[tuple[Fichier, str]] = []
        for perdant in perdants:
            nom_propose = f"{perdant.tag}_{perdant.chemin.name}"
            nom_final = nom_destination_unique(dossier_groupe, nom_propose)
            cible = dossier_groupe / nom_final
            origine = perdant.chemin
            try:
                shutil.move(str(origine), str(cible))
                perdants_deplaces.append((perdant, nom_final))
                deplacements_effectues.append((origine, cible))
                fichiers_deplaces_total += 1
                espace_concerne += perdant.taille
            except (OSError, shutil.Error) as e:
                erreurs.append(f"Déplacement échoué : {origine} → {cible} ({e})")

        groupes_traites.append((index, groupe, gagnant, perdants_deplaces))

        # Réécriture du rapport à chaque tour
        ecrire_rapport_final(
            chemin_rapport,
            chemin_dl,
            chemin_desk,
            fichiers_scannes_total,
            groupes_traites,
            erreurs,
            fichiers_deplaces_total,
            espace_concerne,
            deplacements_effectues,
            interrompu=False,
            debut_scan=debut_scan,
            nb_groupes_prevus=nb_groupes_prevus,
        )

        if DELAI_PROGRESSIF > 0:
            delai = DELAI_PROGRESSIF if index <= SEUIL_BASCULE_RAPIDE else DELAI_PROGRESSIF_RAPIDE
            time.sleep(delai)

    # ───── Phase 4 : finalisation

    if interrompu:
        # Réécriture finale avec le bandeau d'interruption
        ecrire_rapport_final(
            chemin_rapport,
            chemin_dl,
            chemin_desk,
            fichiers_scannes_total,
            groupes_traites,
            erreurs,
            fichiers_deplaces_total,
            espace_concerne,
            deplacements_effectues,
            interrompu=True,
            debut_scan=debut_scan,
            nb_groupes_prevus=nb_groupes_prevus,
        )
        return

    if not groupes_traites:
        # Aucun doublon : on supprime le dossier sortie (vide) et on pose
        # un fichier explicite directement sur le Bureau.
        try:
            chemin_rapport.unlink(missing_ok=True)
            dossier_sortie.rmdir()
        except OSError:
            pass
        message = (
            f"Scan terminé le {datetime.now():%Y-%m-%d}.\n"
            f"Aucun doublon détecté parmi les {fichiers_scannes_total} "
            f"fichiers scannés.\n"
            "Tes dossiers Downloads et Desktop sont propres.\n"
        )
        chemin_aucun = desktop_pour_sortie / "_AUCUN_DOUBLON_TROUVE.txt"
        chemin_aucun.write_text(message, encoding="utf-8")
        ouvrir_explorateur(desktop_pour_sortie)
        return

    # Sinon, le rapport final a déjà été écrit dans la dernière itération
    # de la boucle. Rien à rouvrir.


if __name__ == "__main__":
    # Filet de sécurité global : pythonw n'a pas de stdout, donc une
    # exception non gérée serait invisible. On la matérialise en fichier.
    try:
        main()
    except SystemExit:
        raise
    except BaseException:
        tb = traceback.format_exc()
        try:
            desktop = _desktop_pour_signaux()
            desktop.mkdir(parents=True, exist_ok=True)
            chemin_err = desktop / "_ERREUR_YAPADEUX.txt"
            chemin_err.write_text(
                "yapadeux a planté. Voici la trace technique :\n\n" + tb,
                encoding="utf-8",
            )
            # On tente de libérer le verrou même en cas de crash
            try:
                liberer_verrou(desktop / NOM_VERROU_RUN)
            except OSError:
                pass
            if not MODE_TEST:
                try:
                    os.startfile(str(desktop))  # type: ignore[attr-defined]
                except OSError:
                    pass
        except OSError:
            pass
        sys.exit(1)
