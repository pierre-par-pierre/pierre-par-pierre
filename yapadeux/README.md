# yapadeux — y'a pas deux

> Détecte les fichiers doublons dans tes dossiers Téléchargements et Bureau.
> Pierre par pierre.

## La friction

Mes dossiers Downloads et Desktop accumulent des copies du même fichier
(rapports téléchargés deux fois, photos sauvées en double, brouillons
recopiés). Je veux les retrouver pour les supprimer, sans terminal,
sans ligne de commande, juste un double-clic.

## La stack

Python 3.10+ stdlib uniquement (`hashlib`, `os`, `shutil`, `pathlib`,
`collections`, `datetime`, `time`, `ctypes`) + 3 wrappers `.vbs` qui
appellent `pythonw.exe` pour qu'aucune fenêtre noire n'apparaisse.

## Installation (à faire une seule fois)

1. Télécharge Python depuis https://python.org
2. **Pendant l'install, coche "Add Python to PATH"**
3. C'est tout. `pythonw.exe` (la version sans console) est livré
   automatiquement avec toute install Python sur Windows — c'est ce
   qui permet le mode silencieux.

## Utilisation

### Pour trier les doublons

Double-clique sur **`LANCER.vbs`**. Un dossier `_doublons-detectes/`
apparaît sur ton Bureau et se remplit progressivement. Tu vérifies le
contenu, tu supprimes manuellement ce dont tu ne veux plus.

### Pour arrêter un tri en cours

Double-clique sur **`STOP.vbs`**. Une popup confirme la demande
d'arrêt. Le tri s'arrête au prochain groupe traité (quelques
secondes max). Les fichiers déjà déplacés restent dans
`_doublons-detectes/`.

### Pour tout remettre à sa place

Double-clique sur **`RESTAURER.vbs`**. Tous les fichiers déplacés par
le dernier tri reviennent à leur emplacement d'origine. Si tu en as
déjà supprimé certains, ils restent supprimés (ton choix). Si tu en
as remis certains à la main, le script les ignore proprement (pas
d'écrasement). Un `_RAPPORT_RESTAURATION.txt` apparaît sur le Bureau.

> Si l'Explorateur ne se rafraîchit pas tout seul, appuie sur F5.

## Comment ça marche

1. L'outil scanne ton Downloads en profondeur (sous-dossiers inclus)
   et le Bureau en surface (racine uniquement, pour ne pas toucher à
   tes dossiers de projets et backups).
2. Hash MD5 des fichiers de même taille (octet-bit-identique).
3. Pour chaque groupe de doublons : le fichier le plus ancien du
   Bureau gagne et reste à sa place ; sinon le plus ancien du
   Downloads. Tous les autres sont rangés dans `_doublons-detectes/`.

## Sécurité

yapadeux ne touche QUE des chemins contenant l'un des 4 mots-clés
suivants : `Downloads`, `Téléchargements`, `Desktop`, `Bureau`.
Tout autre chemin est refusé.

Aucune suppression : uniquement du déplacement (`shutil.move`).
Si un fichier déplacé devait être un faux positif, il reste
récupérable dans `_doublons-detectes/` ou via `RESTAURER.vbs`.

En cas de plantage, un fichier `_ERREUR_YAPADEUX.txt` (ou
`_ERREUR_RESTAURATION.txt`) apparaît sur ton Bureau avec la trace
technique.

## Limitations

- **Le scan du Bureau ne descend PAS dans les sous-dossiers.** Si tu
  veux qu'un fichier sur le Bureau soit pris en compte, il doit être
  à la racine du Bureau.
- **Détection MD5 binaire stricte.** Deux PDF visuellement
  identiques avec des métadonnées différentes ne sont **pas**
  détectés comme doublons.
- **Photos compressées différemment** (JPEG 90 % vs 95 %) ne sont
  pas détectées comme doublons.
- **Très gros fichiers** (plusieurs Go) : le hash prend du temps,
  c'est normal. `_RAPPORT.txt` est mis à jour pendant le scan, tu
  peux le rafraîchir (F5) pour suivre la progression.

## Comportement aux re-lancements

- **Premier scan** : crée `_doublons-detectes/` sur le Bureau.
- **Deuxième scan sans restauration entre les deux** : crée
  `_doublons-detectes_<timestamp>/` à côté du précédent. Pas
  d'écrasement.
- **`RESTAURER.vbs` cible TOUJOURS le dossier de tri le plus récent**
  (date de modification max). Pour restaurer un tri plus ancien,
  supprime manuellement le plus récent d'abord.
- **Deux scans simultanés : interdits** par le verrou. Le second
  affiche une popup et quitte.
- **Deux restaurations simultanées : interdites** par le verrou
  équivalent.

## Licence

MIT — voir [LICENSE](LICENSE).
