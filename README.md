# Suivis · Jeff

App web de suivi des engagements (style Trello, en plus fun), branchée sur le Google Sheet « Suivis Jeff ».

## Ce que ça fait
- **Focus** : les 5 dossiers qui comptent, calculés selon la priorité, le retard, le type et l'âge.
- **Tableau** : colonnes par statut, glisser-déposer des cartes.
- **Liste** : regroupée par dossier.
- **Swipe à droite** pour fermer (confettis, son, vibration, +10 ⚡). **Swipe à gauche** pour reporter de 2 jours ouvrables. Bouton **Annuler** sur chaque action.
- **Combos** quand tu fermes plusieurs dossiers de suite, des **niveaux**, une **série** de jours 🔥, un anneau « fermés aujourd'hui ».
- **Ajout rapide en langage naturel** : `Relancer Simon vendredi !haute @Rossana #THD`
  - `!haute` `!moyenne` `!basse` · `@personne` · `#dossier` · `demain`, `vendredi`, `dans 3 jours`, `20 oct.`
  - mots-clés de type : `balle` (on attend après moi), `promis`, `doit`
- Filtre par personne (Cindy voit sa liste), recherche, mode clair/sombre.
- Synchro aux 60 secondes : quand Cindy coche un item, la carte flashe chez toi.
- Hors ligne : les changements sont gardés et envoyés au retour.
- Raccourcis : `N` nouveau, `/` recherche, `1` `2` `3` vues, `Échap` ferme.

## Installation (15 min, une seule fois)

### 1. Optimiser le Google Sheet
1. Ouvre le Sheet « Suivis Jeff » > **Extensions > Apps Script**.
2. Colle tout le contenu de `Code.gs`, puis enregistre.
3. Choisis la fonction `setup` et clique **Exécuter**. Autorise l'accès.
   - Ça crée les onglets `Suivis`, `Listes`, `Journal`, `Config` et `Utilisateurs`, les cases à cocher, les listes déroulantes et les couleurs, puis migre tes 15 suivis avec des identifiants `SUI-0001…`.
   - L'ancien onglet est gardé, caché, sous le nom « Ancien format (archive) ».

### 2. Créer les accès
1. Recharge le Sheet : un menu **🔐 Suivis** apparaît.
2. **🔐 Suivis > Définir un mot de passe** : tape le nom (ex. `Cindy`), puis un mot de passe temporaire (10 caractères minimum, lettres et au moins un chiffre).
3. Donne-le à la personne de vive voix ou par texto. Elle le change elle-même dans l'app (⚙️ > Changer mon mot de passe).
4. Pour ajouter quelqu'un : une ligne dans l'onglet `Utilisateurs` (nom, courriel, rôle, actif coché), puis l'étape 2.
5. Pour couper un accès **immédiatement** : décoche `actif` dans `Utilisateurs`. Sa session meurt à la requête suivante.

### 3. Publier l'API
1. Dans Apps Script : **Déployer > Nouveau déploiement > Application Web**.
2. Exécuter en tant que : **Moi**. Accès : **Tout le monde** (c'est le mot de passe qui protège, pas l'URL).
3. Copie l'URL qui finit par `/exec`.

### 4. Mettre l'app sur GitHub Pages
1. Dans `index.html`, colle l'URL à la ligne `const API_URL = '';`.
2. Dépose `index.html` dans un repo (ex. `jefftremblay-bot/suivis`), puis **Settings > Pages** > Branch `main` / root.
3. Envoie le lien à l'équipe : chacun se connecte avec son nom (ou courriel) et son mot de passe.

## Sécurité
- **Aucun secret dans GitHub.** Le code public ne contient que l'URL de l'API, qui ne répond à rien sans connexion.
- **Mots de passe hachés** (sel aléatoire + SHA-256 itéré), jamais stockés en clair. Les colonnes sel/hash sont cachées et l'onglet `Utilisateurs` est protégé (admin seulement).
- **Sessions** : jeton aléatoire de plus de 350 bits, valide 30 jours (ou jusqu'à la fermeture du navigateur si « Rester connecté » est décoché). Toujours envoyé en POST, jamais dans une URL.
- **Anti-force brute** : 5 essais ratés = compte bloqué 15 minutes, avec délai sur chaque échec.
- **Traçabilité fiable** : le nom inscrit dans `maj_par` et le `Journal` vient de la session côté serveur, pas du navigateur. Les connexions sont aussi journalisées.
- **Révocation** : changer un mot de passe ferme les autres appareils de la personne. Menu **🔐 Suivis > Déconnecter tout le monde** en cas de doute.
- **Anti-injection** : une valeur qui commence par `=`, `+`, `-` ou `@` est enregistrée comme texte, jamais comme formule.
- Limite à connaître : quiconque a un accès d'édition au Google Sheet lui-même voit toutes les données. Garde le partage du Sheet aux personnes de confiance; l'équipe peut passer uniquement par l'app.

## Le Sheet reste la source de vérité
- Cocher la case `fait` dans le Sheet met le statut à « Fait » (et l'inverse), avec la date de fermeture.
- Chaque modification passe par l'onglet `Journal` (qui, quoi, quand).
- Pour ajouter un dossier ou une personne aux menus : onglet `Listes`.
- Le Registre des engagements de Claude (lundi et jeudi, 7h40) lit et alimente ce même onglet `Suivis`.

## Mettre à jour le code plus tard
- Après une modification de `Code.gs` : **Déployer > Gérer les déploiements > Modifier > Nouvelle version**. L'URL ne change pas.
