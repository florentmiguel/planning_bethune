# Planning Béthune

Page web pour téléphone qui sert de pointeuse au tableau de présence `Planning_Bethune_2026.xlsx` (Florent, Gugus, Virginie, du 27/11 au 28/12/2026).

- **Pointer** : arrivée et départ en un geste, à l'heure du téléphone. Un poste de nuit est rattaché au jour d'arrivée.
- **Planning** : correction des heures, signature (texte « ✓ Prénom jj/mm hh:mm »).
- **Totaux & Excel** : totaux par semaine, par mois et général, puis téléchargement du fichier Excel d'origine rempli (mise en forme, formules et tableau conservés).

Hébergement : GitHub Pages (gratuit). Base de données : Supabase, PostgreSQL (offre gratuite).

## Mise en place (environ 15 min, une seule fois)

### 1. Base de données Supabase
1. Crée un compte sur https://supabase.com, puis un projet. Choisis la région Europe (Paris ou Francfort).
2. Va dans **SQL Editor**, puis **New query**. Colle le contenu de `supabase/schema.sql`.
3. **Remplace les 4 codes `CODE-…-A-CHANGER`** (en bas du fichier) par un code personnel pour chacun (8 caractères ou plus), puis lance **Run**.
4. Dans **Project Settings**, puis **API**, copie le **Project URL** et la clé **anon public**.

### 2. Configuration
Ouvre `config.js` et colle l'URL et la clé anon à la place de `VOTRE-PROJET` et `VOTRE-CLE-ANON`.

Pour changer un prénom (par exemple remplacer « Gugus » après le recrutement), modifie la liste `prenoms` dans `config.js`. Les pointages déjà faits sont conservés.

### 3. Mise en ligne sur GitHub Pages
1. Pousse ce dossier dans un dépôt GitHub.
   Nomme le dépôt **`planning-bethune`** : l'adresse sera alors `https://<ton-compte>.github.io/planning-bethune/`.
2. Dans le dépôt, va dans **Settings**, puis **Pages**. Choisis **Source : Deploy from a branch**, la branche **main**, le dossier **/ (root)**.
3. L'adresse s'affiche. Pour avoir exactement `https://planning-bethune.github.io`, crée plutôt un compte (ou une organisation) GitHub nommé `planning-bethune` et un dépôt nommé `planning-bethune.github.io`.

### 4. Sur les téléphones
Chacun ouvre l'adresse et saisit son code personnel une fois (il est ensuite mémorisé). Ajoute la page à l'écran d'accueil (sur iPhone : Partager, puis « Sur l'écran d'accueil ») pour l'avoir comme une appli.

## Accès (un code personnel par personne)
| Personne | Rôle | Voit | Modifie | Excel |
|---|---|---|---|---|
| Florent | admin | tout | tout | oui |
| Virginie | admin | tout | tout | oui |
| Anne-Claire (employeur) | admin | tout | tout | oui |
| Poste en recrutement | salarie | son planning | son planning | non |

Le rôle `responsable` existe aussi : il voit tout et peut télécharger l'Excel, mais ne modifie que sa propre colonne.

**Arrivée de la nouvelle recrue.** Dans Supabase, rubrique SQL Editor :
```sql
update acces set actif = false where nom = 'Gugus';
insert into acces (code, nom, role, col) values ('SON-CODE', 'Prénom', 'salarie', 'F');
```
Remplace aussi « Gugus » dans `prenoms` (fichier `config.js`). Ses pointages restent dans la colonne F.

**Retirer un accès :** `update acces set actif = false where nom = '…';`

## Sécurité
- Les tables sont fermées : la clé publique ne donne accès qu'aux fonctions `ptg_me`, `ptg_list` et `ptg_upsert`, qui vérifient le code et les droits côté base. Un salarié ne peut donc ni lire ni modifier le planning des autres, même en bricolant la page.
- Toute modification est tracée avec son auteur dans `pointages_log`.

## Hors connexion
Les pointages faits sans réseau restent sur le téléphone et sont envoyés automatiquement au retour du réseau.

## Nouvelle période
Le modèle Excel est dans `planning.js` (généré depuis le `.xlsx`). Pour un nouveau planning, il faut régénérer ce fichier à partir du nouveau tableau.
