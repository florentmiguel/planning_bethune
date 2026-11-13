# Planning Événements — pointeuse mobile

Pointage, signature, totaux et export Excel pour **plusieurs événements en parallèle** (marché de Noël, salons…), chacun avec ses dates et **son équipe**.

Hébergement : GitHub Pages. Base de données : Supabase (PostgreSQL).

## Fichiers
| Fichier | Rôle |
|---|---|
| `index.html` | la page |
| `xlsx.js` | génération du fichier Excel (planning complet d'un événement ou détail d'une personne) |
| `config.js` | adresse et clé **Publishable** Supabase |
| `supabase/schema.sql` | installation neuve (tables + fonctions + premiers comptes + premier événement) |
| `supabase/migration-v3.sql` | passage depuis la version « colonnes B/F/J » — **une seule fois** |
| `supabase/fonctions.sql` | fonctions seules (à relancer après une mise à jour des fonctions) |

## Installation neuve
1. Supabase > SQL Editor : coller `supabase/schema.sql`, remplacer les 4 codes `CODE-…-A-CHANGER`, puis **Run**.
2. Renseigner `config.js` : Project URL (`https://xxxx.supabase.co`) et clé **Publishable** (`sb_publishable_…`).
3. Envoyer le dossier sur GitHub, puis Settings > Pages > branche `main`, dossier `/ (root)`.

## Mise à jour depuis la version « colonnes B/F/J »
Supabase > SQL Editor : coller `supabase/migration-v3.sql` > **Run** (une seule fois). Comptes, codes et pointages sont conservés ; l'événement « Marché de Noël de Béthune 2026 » est créé avec l'équipe actuelle. Remplacer ensuite `index.html` et `xlsx.js` sur GitHub et supprimer `planning.js`.

## Droits (rôle global, valable pour tous les événements)
| Rôle | Événements | Voit | Pointe / signe | Corrige les heures | Administration |
|---|---|---|---|---|---|
| admin | tous | tout le monde | tout le monde | tout le monde | comptes, événements, équipes, Excel complet |
| responsable | ceux de ses équipes | lui + les salariés | lui + les salariés | lui + les salariés | non |
| salarié | ceux de ses équipes | lui seul | lui seul, **heure du serveur** | non | non |

Chacun peut exporter **son détail** (ou celui d'une personne qu'il voit) en .xlsx.

## Administration (onglet Totaux, administrateur)
- **Événements** : nom, date de début, date de fin (400 jours max.), équipe (cases à cocher) et **place** de chacun (ordre des colonnes dans l'Excel), archivé.
- **Comptes** : prénom, code (8 caractères min.), rôle, actif.
- Les heures sont rattachées à la personne **et** à l'événement : une même personne peut travailler sur deux événements, ses heures sont séparées.
- Si plusieurs événements sont visibles, un menu en haut de la page permet de choisir l'événement.

## Onglets
| Onglet | Qui | Rôle |
|---|---|---|
| Pointer | tous | pointage personnel (l'admin peut pointer pour n'importe qui) ; affiche l'horaire prévu |
| Pointages | responsables, admin | corriger les heures réellement pointées |
| Horaires | admin | saisir les horaires **prévus** (jusqu'à 2 créneaux par jour) |
| Totaux | tous | totaux, détail dépliant, exports Excel |
| Admin | admin | événements, équipes, comptes |

## Excel
Généré par l'appli : une ligne par jour, un bloc Début / Fin / Durée / Signature par personne, totaux par semaine, par mois et général (formules Excel), en-têtes figés, impression A4 paysage sur la largeur d'une page.

## Sécurité
Tables fermées (RLS sans politique) ; la clé publique ne donne accès qu'aux fonctions `ptg_*`, qui vérifient le code, l'événement et les droits. Toute modification est tracée avec son auteur dans `pointages_log`.
