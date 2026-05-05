"""
restaurer.py — annule le dernier tri yapadeux.

Lit la section MAPPING du _RAPPORT.txt du dossier de tri le plus récent
sur le Desktop, et remet chaque fichier déplacé à son emplacement
d'origine via shutil.move (jamais d'écrasement, jamais de suppression
destructive).

Lancement prévu : double-clic sur RESTAURER.vbs.

Cas gérés :
- Aucun dossier _doublons-detectes/ → _RIEN_A_RESTAURER.txt + popup.
- Fichier supprimé manuellement entre temps → "Déjà absent", on continue.
- Origine déjà ré-occupée → "Conflit", fichier laissé dans le dossier
  de tri, on continue.
- Si le dossier _doublons-detectes/ est vide après restauration, on le
  supprime. Sinon on le laisse pour traitement manuel.
"""

import ctypes
import os
import shutil
import sys
import traceback
from datetime import datetime
from pathlib import Path


# ─────────────────────────────────────────────────────────────────────
# Constantes
# ─────────────────────────────────────────────────────────────────────

MODE_TEST = False
CHEMIN_DESKTOP_TEST: Path | None = None

NOM_VERROU_RESTAURATION = "_yapadeux-restoring.lock"
PREFIXE_DOSSIER_TRI = "_doublons-detectes"

MARQUEUR_MAPPING_DEBUT = "MAPPING:"
MARQUEUR_MAPPING_FIN = "END_MAPPING"
SEPARATEUR_MAPPING = "|"

LIGNE = "─" * 39
LIGNE_DOUBLE = "═" * 39

# Mémorisation de la dernière popup pour les tests
derniere_popup: tuple[str, str] | None = None

# Constantes Win32
_MB_OK = 0x0
_MB_ICONINFORMATION = 0x40
_MB_ICONWARNING = 0x30
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_STILL_ACTIVE = 259


# ─────────────────────────────────────────────────────────────────────
# Helpers Windows
# ─────────────────────────────────────────────────────────────────────

def afficher_popup(titre: str, message: str, icone: int = _MB_ICONINFORMATION):
    """MessageBoxW. En mode test : mémorise dans derniere_popup."""
    global derniere_popup
    derniere_popup = (titre, message)
    if MODE_TEST:
        return
    try:
        ctypes.windll.user32.MessageBoxW(0, message, titre, _MB_OK | icone)
    except OSError:
        pass


def processus_actif(pid: int) -> bool:
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
    if verrou.exists():
        try:
            pid_existant = int(verrou.read_text(encoding="utf-8").strip())
        except (ValueError, OSError):
            pid_existant = 0
        if pid_existant and processus_actif(pid_existant):
            afficher_popup(
                "yapadeux — Restauration",
                "Une restauration est déjà en cours.\n\n"
                "Attends qu'elle se termine avant d'en lancer une autre.",
                icone=_MB_ICONWARNING,
            )
            return False
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
    try:
        if verrou.exists():
            verrou.unlink()
    except OSError:
        pass


def ouvrir_explorateur(chemin: Path):
    if MODE_TEST:
        return
    try:
        os.startfile(str(chemin))  # type: ignore[attr-defined]
    except OSError:
        pass


# ─────────────────────────────────────────────────────────────────────
# Détection du dossier de tri à restaurer
# ─────────────────────────────────────────────────────────────────────

def desktop_courant() -> Path:
    """Desktop utilisé. En mode test, faux-Desktop."""
    if MODE_TEST and CHEMIN_DESKTOP_TEST is not None:
        return CHEMIN_DESKTOP_TEST
    home = Path.home()
    for nom in ("Desktop", "Bureau"):
        p = home / nom
        if p.is_dir():
            return p
    return home / "Desktop"


def trouver_dossier_tri_recent(desktop: Path) -> Path | None:
    """
    Liste les dossiers du Desktop dont le nom commence par
    `_doublons-detectes`, retourne le plus récent (mtime max).
    """
    candidats = []
    try:
        for entree in desktop.iterdir():
            if entree.is_dir() and entree.name.startswith(PREFIXE_DOSSIER_TRI):
                candidats.append(entree)
    except OSError:
        return None
    if not candidats:
        return None
    candidats.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidats[0]


# ─────────────────────────────────────────────────────────────────────
# Parsing du MAPPING
# ─────────────────────────────────────────────────────────────────────

def parser_mapping(rapport: Path) -> list[tuple[Path, Path]]:
    """
    Lit le _RAPPORT.txt et extrait les couples (origine, destination)
    entre les marqueurs MAPPING: et END_MAPPING.
    """
    couples: list[tuple[Path, Path]] = []
    try:
        contenu = rapport.read_text(encoding="utf-8")
    except OSError:
        return couples

    lignes = contenu.splitlines()
    dans_mapping = False
    for ligne in lignes:
        ligne_strippee = ligne.strip()
        if ligne_strippee == MARQUEUR_MAPPING_DEBUT:
            dans_mapping = True
            continue
        if ligne_strippee == MARQUEUR_MAPPING_FIN:
            dans_mapping = False
            continue
        if dans_mapping and SEPARATEUR_MAPPING in ligne_strippee:
            origine_str, _, destination_str = ligne_strippee.partition(SEPARATEUR_MAPPING)
            if origine_str and destination_str:
                couples.append((Path(origine_str), Path(destination_str)))
    return couples


# ─────────────────────────────────────────────────────────────────────
# Restauration (le coeur)
# ─────────────────────────────────────────────────────────────────────

def restaurer_un(origine: Path, destination: Path) -> str:
    """
    Tente de remettre `destination` à l'emplacement `origine`.
    Retourne :
      - "restaure"  : déplacement OK
      - "absent"    : destination n'existe plus (supprimée manuellement)
      - "conflit"   : origine déjà occupée
      - "erreur"    : échec OSError
    """
    if not destination.exists():
        return "absent"
    if origine.exists():
        return "conflit"
    try:
        origine.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(destination), str(origine))
        return "restaure"
    except (OSError, shutil.Error):
        return "erreur"


def ecrire_rapport_restauration(
    chemin_rapport: Path,
    nom_dossier_source: str,
    debut: datetime,
    resultats: list[tuple[Path, Path, str]],
    dossier_supprime: bool,
    nb_conflits: int,
):
    """Écrit _RAPPORT_RESTAURATION.txt sur le Desktop."""
    nb_restaures = sum(1 for _, _, s in resultats if s == "restaure")
    nb_absents = sum(1 for _, _, s in resultats if s == "absent")
    nb_erreurs = sum(1 for _, _, s in resultats if s == "erreur")

    lignes: list[str] = []
    lignes.append(LIGNE_DOUBLE)
    lignes.append("   YAPADEUX — Restauration")
    lignes.append(LIGNE_DOUBLE)
    lignes.append("")
    lignes.append(f"Date          : {debut:%Y-%m-%d %H:%M}")
    lignes.append(f"Source        : {nom_dossier_source}/")
    lignes.append("")
    lignes.append(LIGNE)
    lignes.append("Bilan")
    lignes.append(LIGNE)
    lignes.append(f"Restaurés     : {nb_restaures}")
    lignes.append(
        f"Déjà absents  : {nb_absents}   "
        "(fichiers supprimés manuellement avant la restauration)"
    )
    lignes.append(
        f"Conflits      : {nb_conflits}   "
        "(fichiers laissés dans _doublons-detectes/)"
    )
    if nb_erreurs:
        lignes.append(f"Erreurs       : {nb_erreurs}   (déplacement impossible)")
    lignes.append("")
    lignes.append(LIGNE)
    lignes.append("Détail")
    lignes.append(LIGNE)
    if not resultats:
        lignes.append("Aucun fichier à restaurer (mapping vide).")
    for origine, destination, statut in resultats:
        if statut == "restaure":
            lignes.append(f"✓ Restauré : {origine}")
        elif statut == "absent":
            lignes.append(
                f"✗ Déjà absent : {origine} "
                "(le fichier n'était plus dans le dossier de tri)"
            )
        elif statut == "conflit":
            lignes.append(
                f"✗ Conflit : {origine} "
                "(un fichier porte déjà ce nom à l'origine)"
            )
        else:
            lignes.append(f"✗ Erreur : {origine} → {destination}")
    lignes.append("")
    lignes.append(LIGNE)
    lignes.append("Que faire maintenant ?")
    lignes.append(LIGNE)
    if nb_conflits > 0:
        lignes.append(
            f"Le dossier _doublons-detectes/ contient encore "
            f"{nb_conflits} fichier(s) en conflit."
        )
        lignes.append(
            "Ouvre-le pour décider manuellement (renommer, déplacer, supprimer)."
        )
    elif dossier_supprime:
        lignes.append("Tous les fichiers ont été remis à leur place.")
        lignes.append("Le dossier _doublons-detectes/ a été supprimé.")
    else:
        lignes.append(
            "Le dossier _doublons-detectes/ contient encore des fichiers "
            "(erreurs ou résidus)."
        )
        lignes.append("Ouvre-le pour vérifier ce qui reste.")
    lignes.append("")

    chemin_rapport.write_text("\n".join(lignes), encoding="utf-8")


def dossier_est_vide_de_fichiers(racine: Path) -> bool:
    """True si plus aucun fichier dans la sous-arborescence (sauf _RAPPORT.txt
    qu'on ignore — on le supprimera nous-même)."""
    for chemin in racine.rglob("*"):
        if chemin.is_file() and chemin.name != "_RAPPORT.txt":
            return False
    return True


def supprimer_dossier_tri(racine: Path):
    """Supprime le dossier de tri (devrait ne contenir que _RAPPORT.txt)."""
    rapport = racine / "_RAPPORT.txt"
    try:
        if rapport.exists():
            rapport.unlink()
    except OSError:
        pass
    # Suppression récursive des sous-dossiers vides
    for sous in sorted(racine.rglob("*"), reverse=True):
        try:
            if sous.is_dir():
                sous.rmdir()
        except OSError:
            pass
    try:
        racine.rmdir()
    except OSError:
        pass


# ─────────────────────────────────────────────────────────────────────
# main
# ─────────────────────────────────────────────────────────────────────

def main():
    debut = datetime.now()
    desktop = desktop_courant()
    desktop.mkdir(parents=True, exist_ok=True)

    verrou = desktop / NOM_VERROU_RESTAURATION
    if not acquerir_verrou(verrou):
        sys.exit(1)

    try:
        dossier_tri = trouver_dossier_tri_recent(desktop)
        if dossier_tri is None:
            # Rien à restaurer : popup + fichier explicite
            chemin_rien = desktop / "_RIEN_A_RESTAURER.txt"
            chemin_rien.write_text(
                f"Vérification le {datetime.now():%Y-%m-%d %H:%M}.\n"
                "Aucun dossier _doublons-detectes/ trouvé sur le Bureau.\n"
                "Il n'y a rien à restaurer.\n",
                encoding="utf-8",
            )
            afficher_popup(
                "yapadeux — Restauration",
                "Rien à restaurer.\n\n"
                "Aucun dossier _doublons-detectes/ trouvé sur ton Bureau.",
            )
            ouvrir_explorateur(desktop)
            return

        rapport_source = dossier_tri / "_RAPPORT.txt"
        couples = parser_mapping(rapport_source)

        resultats: list[tuple[Path, Path, str]] = []
        for origine, destination in couples:
            statut = restaurer_un(origine, destination)
            resultats.append((origine, destination, statut))

        nb_conflits = sum(1 for _, _, s in resultats if s == "conflit")

        # On peut supprimer le dossier de tri si plus aucun fichier
        # restant (les fichiers en conflit y sont encore).
        dossier_supprime = False
        if dossier_est_vide_de_fichiers(dossier_tri):
            supprimer_dossier_tri(dossier_tri)
            dossier_supprime = True

        # Écriture du rapport de restauration sur le Desktop
        chemin_rapport_restau = desktop / "_RAPPORT_RESTAURATION.txt"
        ecrire_rapport_restauration(
            chemin_rapport_restau,
            nom_dossier_source=dossier_tri.name,
            debut=debut,
            resultats=resultats,
            dossier_supprime=dossier_supprime,
            nb_conflits=nb_conflits,
        )

        ouvrir_explorateur(desktop)
    finally:
        liberer_verrou(verrou)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except BaseException:
        tb = traceback.format_exc()
        try:
            desktop = desktop_courant()
            desktop.mkdir(parents=True, exist_ok=True)
            chemin_err = desktop / "_ERREUR_RESTAURATION.txt"
            chemin_err.write_text(
                "restaurer.py a planté. Voici la trace technique :\n\n" + tb,
                encoding="utf-8",
            )
            try:
                liberer_verrou(desktop / NOM_VERROU_RESTAURATION)
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
