SKILL — Google Workspace Integration BRASUS
Version : 2.0 — Alignée Vision Stratégique Complète
Auteur : Digital Mastery Agency (DMA) — Abidjan, Côte d'Ivoire
Statut : Production-Ready

## Objectif Fondamental

Google Workspace n'est pas le système de gestion de BRASUS.
Google Workspace est l'interface opérationnelle de BRASUS.

Claude = le cerveau.
Google Workspace = les écrans de saisie et de visualisation.

Toutes les données saisies dans Google Workspace sont collectées, analysées et transformées par Claude avant toute consolidation, rapport ou décision. Les utilisateurs ne voient jamais la logique interne. Ils voient uniquement leur espace de travail.

## Architecture de Cloisonnement Départemental

### Principe Absolu : Un Département = Un Fichier = Un Accès

Chaque département dispose d'un fichier Google Sheets entièrement indépendant, protégé par les permissions Google Workspace.

```
Google Drive Client/
├── [ADMIN_BRASUS_ONLY]/
│   ├── BRASUS_CORE_CONFIG.sheets          ← Accès admin DMA uniquement
│   ├── BRASUS_SYNC_ENGINE.sheets          ← Accès admin DMA uniquement
│   └── BRASUS_AGENTS_LOG.sheets           ← Accès admin DMA uniquement
│
├── DEPARTEMENTS/
│   ├── COMMERCIAL.sheets                  ← Responsable Commercial uniquement
│   ├── MARKETING.sheets                   ← Responsable Marketing uniquement
│   ├── RESSOURCES_HUMAINES.sheets         ← Responsable RH uniquement
│   ├── FINANCE.sheets                     ← Responsable Finance uniquement
│   └── OPERATIONS.sheets                  ← Responsable Opérations uniquement
│
└── DIRECTION_GENERALE/
    └── COCKPIT_DG.sheets                  ← Directeur Général uniquement
```

### Matrice des Permissions Google Workspace

| Rôle | Son département | Autres départements | Admin BRASUS | Cockpit DG | Agents IA |
|------|----------------|---------------------|--------------|------------|-----------|
| Employé | Tâches assignées uniquement | ✗ Aucun accès | ✗ | ✗ | ✗ |
| Responsable | ✅ Lecture + Saisie | ✗ Aucun accès | ✗ | ✗ | ✗ |
| Directeur Général | ✗ | ✗ | ✗ | ✅ Lecture seule | ✗ |
| Administrateur BRASUS | ✅ Complet | ✅ Complet | ✅ Complet | ✅ Complet | ✅ Complet |

### Configuration Technique du Cloisonnement

```javascript
// Paramétrage des permissions par fichier départemental
const permissionsConfig = {
  departement: {
    owner: "admin@dma-brasus.com",          // Propriétaire = DMA uniquement
    editors: ["responsable@client.com"],     // Éditeur = 1 seul responsable
    viewers: [],                             // Aucun viewer externe
    sharedWithLink: false,                  // Partage par lien DÉSACTIVÉ
    publicAccess: false                     // Accès public DÉSACTIVÉ
  }
};

// Révocation automatique si changement de responsable
async function updateDepartmentAccess(oldEmail, newEmail, sheetId) {
  await removePermission(sheetId, oldEmail);
  await addEditor(sheetId, newEmail);
  await logAccessChange(sheetId, oldEmail, newEmail, timestamp());
}
```

## Protection de la Propriété Intellectuelle BRASUS

### Ce que les Responsables Ne Voient JAMAIS

Les couches suivantes restent totalement invisibles aux utilisateurs départementaux :

- Les prompts et instructions des agents Claude
- Les raisonnements et analyses intermédiaires de Claude
- Les recommandations stratégiques globales de l'entreprise
- Les données des autres départements
- Les procédures internes BRASUS
- Les configurations MCP et automatisations
- Les comparaisons inter-départementales
- Les alertes stratégiques destinées à la Direction

### Séparation des Couches

```
COUCHE VISIBLE (Responsables)
─────────────────────────────
│  Missions du jour          │  ← Résultat de l'analyse Claude, pas la logique
│  Zone de saisie            │  ← Input uniquement
│  Mes performances          │  ← Données personnelles uniquement
│  Mon historique            │  ← Données personnelles uniquement
─────────────────────────────

COUCHE INVISIBLE (Admin BRASUS + Claude)
─────────────────────────────────────────
│  Moteur de génération des missions     │
│  Algorithme de calcul des KPIs         │
│  Règles de verrouillage temporel       │
│  Synchronisation vers Cockpit DG       │
│  Agents Claude par département         │
│  Analyses inter-départementales        │
│  Alertes stratégiques                  │
─────────────────────────────────────────
```

**Règle absolue :** Aucune formule Google Sheets ne doit référencer un fichier d'un autre département. Toutes les consolidations sont réalisées par Claude via MCP, pas par des formules cross-fichiers visibles.

## Structure de l'Espace de Travail Départemental

### Vue du Responsable — Ce qu'il voit chaque matin

Le responsable ouvre un seul fichier. Il voit cinq onglets.
Ce n'est pas un tableur. C'est son poste de commandement personnel.

### Onglet 1 — MISSIONS DU JOUR

Généré automatiquement par Claude chaque matin à 7h00

Claude analyse avant de générer :
- Les objectifs mensuels de l'entreprise
- Les objectifs spécifiques du département
- Les performances des jours précédents
- Les tâches non finalisées (carry-forward)
- Les priorités stratégiques du moment
- Les alertes en cours

Structure de l'onglet :

| # | Mission | Priorité | Objectif lié | Résultat attendu | Statut | Commentaire |
|---|---------|----------|--------------|------------------|--------|-------------|
| 1 | [Générée par Claude] | 🔴 Critique | [Objectif] | [Mesurable] | ☐ | |
| 2 | [Générée par Claude] | 🟡 Haute | [Objectif] | [Mesurable] | ☐ | |
| 3 | [Générée par Claude] | 🟢 Standard | [Objectif] | [Mesurable] | ☐ | |

**Règles de génération des missions :**
- Minimum 3, maximum 8 missions par jour
- Chaque mission est liée à un objectif mesurable
- Les missions en retard apparaissent en rouge avec flag "PRIORITÉ"
- Le responsable ne peut pas modifier les missions — seulement les exécuter

**Cellules modifiables par le responsable :**
- Colonne "Statut" : ☐ Non commencé → 🔄 En cours → ✅ Réalisé → ❌ Non réalisé
- Colonne "Commentaire" : texte libre (500 caractères max)

**Cellules verrouillées (non modifiables) :**
- Toutes les autres colonnes — générées et protégées par Claude

### Onglet 2 — EXÉCUTION JOURNALIÈRE

Zone de saisie active — modifiable uniquement le jour J

Le responsable renseigne ses réalisations en temps réel :

```
┌─────────────────────────────────────────────────┐
│  DATE : [Aujourd'hui]     DÉPARTEMENT : [X]      │
├─────────────────────────────────────────────────┤
│  SECTION A — Actions Réalisées                  │
│  [Zone de saisie libre — obligatoire]            │
├─────────────────────────────────────────────────┤
│  SECTION B — Résultats Obtenus                  │
│  [Zone de saisie structurée selon les missions] │
├─────────────────────────────────────────────────┤
│  SECTION C — Difficultés Rencontrées            │
│  [Zone de saisie libre — facultative]            │
├─────────────────────────────────────────────────┤
│  SECTION D — Besoins / Escalades                │
│  [Zone de saisie — déclenche alerte si renseigné]│
└─────────────────────────────────────────────────┘
```

### Onglet 3 — MES PERFORMANCES

Vue personnelle — lecture seule — mise à jour automatique par Claude

```
PERFORMANCES — [Prénom du Responsable]          Mis à jour : [Timestamp]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
AUJOURD'HUI        Taux d'exécution : [X%]      [Barre de progression]
CETTE SEMAINE      Taux d'exécution : [X%]      [Barre de progression]
CE MOIS            Taux d'exécution : [X%]      [Barre de progression]
CETTE ANNÉE        Taux d'exécution : [X%]      [Barre de progression]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
MON OBJECTIF MENSUEL           [Objectif]
MON RÉSULTAT ACTUEL            [Réalisé]
ÉCART                          [+ ou - X%]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
MON PLAN D'AMÉLIORATION (généré par Claude)
→ [Recommandation 1]
→ [Recommandation 2]
→ [Recommandation 3]
```

Les recommandations sont personnalisées par Claude. Elles ne révèlent jamais les données des autres départements ni la logique BRASUS.

### Onglet 4 — MON HISTORIQUE

Consultation uniquement — lecture seule permanente

| Date | Missions assignées | Missions réalisées | Taux | Points forts | Points à améliorer |
|------|-------------------|-------------------|------|--------------|-------------------|
| [J-1] | [N] | [N] | [%] | [Généré Claude] | [Généré Claude] |
| [J-2] | [N] | [N] | [%] | [Généré Claude] | [Généré Claude] |

Accès disponible :
- Historique journalier : 90 derniers jours
- Historique hebdomadaire : 52 dernières semaines
- Historique mensuel : 24 derniers mois

**Règle absolue :** Aucune ligne de l'historique ne peut être modifiée, supprimée ou ajoutée manuellement.

### Onglet 5 — MON PLAN D'AMÉLIORATION

Généré automatiquement par Claude chaque lundi matin

```
PLAN D'AMÉLIORATION — [Semaine du XX au XX]
Préparé par BRASUS pour [Prénom]

DIAGNOSTIC
Votre taux d'exécution cette semaine : [X%]
Tendance : [↑ En progression / → Stable / ↓ En baisse]

VOS 3 AXES D'AMÉLIORATION PRIORITAIRES
1. [Axe 1 — concret et actionnable]
2. [Axe 2 — concret et actionnable]
3. [Axe 3 — concret et actionnable]

OBJECTIF DE LA SEMAINE PROCHAINE
[Objectif spécifique, mesurable, défini par Claude]

RESSOURCES RECOMMANDÉES
[Si applicable : template, guide, ou action spécifique]
```

## Gestion Temporelle Verrouillée

### Cycle de Vie des Données

```
07:00  ─── Génération automatique des missions du jour (Claude)
           ↓
09:00  ─── Ouverture du fichier par le responsable
           ↓
[Journée] ── Saisie des actions, résultats, commentaires (Onglet 2)
           ↓
18:00  ─── Clôture journalière automatique
           │  ├── Horodatage de toutes les cellules (timestamp UTC)
           │  ├── Verrouillage définitif des cellules du jour
           │  ├── Calcul automatique du taux d'exécution du jour
           │  ├── Archivage vers Onglet Historique
           │  └── Envoi des données à Claude pour analyse
           ↓
18:15  ─── Mise à jour des Performances (Onglet 3)
           ↓
[Vendredi 18:30] ── Consolidation hebdomadaire automatique
           ↓
[1er du mois] ── Consolidation mensuelle automatique
           ↓
[31 décembre] ── Consolidation annuelle automatique
```

### Règles de Verrouillage — Non Négociables

```javascript
// Verrouillage automatique à 18h00 heure d'Abidjan (UTC+0)
const LOCK_TIME_UTC = "18:00:00";

async function dailyLockProcess(sheetId, date) {
  // 1. Horodater chaque ligne saisie
  await timestampAllEntries(sheetId, date);
  
  // 2. Verrouiller les cellules de saisie du jour
  await lockRange(sheetId, `EXECUTION_${date}`, {
    editors: [],  // Retire tous les droits d'édition
    reason: `BRASUS_DAILY_LOCK_${date}`
  });
  
  // 3. Copier vers l'historique
  await appendToHistory(sheetId, date);
  
  // 4. Envoyer à Claude pour analyse
  await sendToBRASUS_Core(sheetId, date);
  
  // 5. Logger l'opération
  await logOperation("DAILY_LOCK", sheetId, date, timestamp());
}
```

**Conséquences du verrouillage :**
- Aucune modification possible après 18h00, même par le responsable
- Seul l'Administrateur BRASUS peut déverrouiller une journée (procédure d'exception documentée)
- Toute tentative de modification est loggée et signalée

## Synchronisation vers le Cockpit du Directeur Général

### Architecture de Consolidation

```
[COMMERCIAL.sheets]    ──┐
[MARKETING.sheets]     ──┤
[RH.sheets]            ──┤──→ CLAUDE (Analyse & Consolidation) ──→ COCKPIT_DG.sheets
[FINANCE.sheets]       ──┤
[OPERATIONS.sheets]    ──┘
```

**Important :** La synchronisation ne se fait JAMAIS par formules cross-fichiers.
Elle se fait UNIQUEMENT via Claude + MCP Google Sheets API.
Le DG ne peut pas "remonter" jusqu'aux fichiers départementaux via le Cockpit.

### Structure du COCKPIT_DG.sheets

**Feuille 1 — TABLEAU DE BORD GLOBAL**

```
┌──────────────────────────────────────────────────────────────────┐
│  BRASUS — COCKPIT DIRECTION GÉNÉRALE          [Date] [Heure]     │
│  {{NOM_ENTREPRISE}}                                              │
├──────────────────┬──────────────────┬──────────────────┬─────────┤
│  CA CE MOIS      │  TRÉSORERIE      │  TAUX EXÉCUTION  │ ALERTES │
│  [X FCFA]        │  [X FCFA]        │  GLOBAL [X%]     │  [N]    │
│  vs obj: [±X%]   │  Seuil: [OK/⚠️]  │  Tendance [↑↓→]  │  [🔴🟡] │
└──────────────────┴──────────────────┴──────────────────┴─────────┘
```

**Feuilles 2 à 6 — VUES DÉPARTEMENTALES CONSOLIDÉES**

| Indicateur | Hier | Cette semaine | Ce mois | Objectif | Écart |
|-----------|------|--------------|---------|----------|-------|
| KPI 1 | [x] | [x] | [x] | [x] | [±%] |
| KPI 2 | [x] | [x] | [x] | [x] | [±%] |
| Taux exécution | [%] | [%] | [%] | 100% | [±%] |

**Feuille 7 — RAPPORTS IA (Claude)**

```
ANALYSE CLAUDE — [Date]
━━━━━━━━━━━━━━━━━━━━━━━
SITUATION GLOBALE
[Synthèse en 3 lignes — générée par Claude]

POINTS FORTS DE LA PÉRIODE
→ [Point 1]
→ [Point 2]

POINTS DE VIGILANCE
⚠️ [Alerte 1 avec niveau de criticité]
⚠️ [Alerte 2 avec niveau de criticité]

RECOMMANDATIONS STRATÉGIQUES
1. [Action prioritaire — délai — responsable]
2. [Action prioritaire — délai — responsable]
3. [Action prioritaire — délai — responsable]

INDICATEURS À SURVEILLER CETTE SEMAINE
▸ [KPI 1] → Seuil cible : [X]
▸ [KPI 2] → Seuil cible : [X]
▸ [KPI 3] → Seuil cible : [X]
```

**Feuille 8 — PLAN D'ACTIONS PRIORITAIRES**

Généré automatiquement par Claude chaque lundi matin :

| # | Action | Département concerné | Urgence | Échéance | Statut |
|---|--------|---------------------|---------|----------|--------|
| 1 | [Action] | [Dept] | 🔴 | [Date] | ☐ |
| 2 | [Action] | [Dept] | 🟡 | [Date] | ☐ |

## Suivi des Performances Individuelles

### Calcul Automatique des Taux d'Exécution

```
FORMULE TAUX D'EXÉCUTION JOURNALIER
────────────────────────────────────
= (Missions réalisées ✅ / Total missions assignées) × 100

FORMULE TAUX D'EXÉCUTION HEBDOMADAIRE
──────────────────────────────────────
= Moyenne des taux journaliers de la semaine
  Pondéré par le nombre de missions par jour

FORMULE TAUX D'EXÉCUTION MENSUEL
─────────────────────────────────
= Moyenne des taux hebdomadaires du mois
  Excluant les jours fériés et congés validés

FORMULE TAUX D'EXÉCUTION ANNUEL
─────────────────────────────────
= Moyenne des taux mensuels de l'année
```

### Niveaux de Performance (Interprétation Claude)

| Taux | Niveau | Couleur | Action Claude |
|------|--------|---------|---------------|
| 90–100% | Excellent | 🟢 Vert | Message de reconnaissance |
| 75–89% | Bon | 🔵 Bleu | Encouragement + 1 conseil |
| 60–74% | Moyen | 🟡 Jaune | Plan d'amélioration ciblé |
| 40–59% | Insuffisant | 🟠 Orange | Alerte RH + plan de soutien |
| < 40% | Critique | 🔴 Rouge | Alerte Direction + escalade |

Les seuils de niveau sont configurables par client dans `/config/client-profile.md`

## Rôle de Claude dans le Flux Google Workspace

### Claude comme Seul Cerveau du Système

```
DONNÉES BRUTES (Google Sheets)
         ↓
    [Collecte MCP]
         ↓
   CLAUDE ANALYSE
   ┌─────────────────────────────────────┐
   │ • Calcule les taux d'exécution      │
   │ • Identifie les tendances           │
   │ • Détecte les anomalies             │
   │ • Compare vs objectifs              │
   │ • Génère les recommandations        │
   │ • Produit les alertes               │
   │ • Prépare les missions du lendemain │
   └─────────────────────────────────────┘
         ↓
RÉSULTATS STRUCTURÉS (écrits dans les Sheets)
   ├── Performances individuelles (Onglet 3)
   ├── Plans d'amélioration (Onglet 5)
   ├── Cockpit DG (toutes les feuilles)
   ├── Rapports envoyés par Gmail
   └── Notifications WhatsApp Business
```

### Planification des Opérations Claude Automatiques

| Heure (UTC+0) | Opération Claude | Fichiers concernés |
|--------------|-----------------|-------------------|
| 07:00 | Génération des missions du jour | Tous les Sheets département |
| 12:00 | Vérification mi-journée + alertes si blocage | Tous les Sheets |
| 18:00 | Clôture + verrouillage + collecte données | Tous les Sheets département |
| 18:15 | Mise à jour performances individuelles | Onglet 3 de chaque département |
| 18:30 | Consolidation et mise à jour Cockpit DG | COCKPIT_DG.sheets |
| 18:45 | Envoi rapport journalier (Gmail + WhatsApp) | DG uniquement |
| Lundi 07:00 | Génération plan d'amélioration hebdo | Onglet 5 de chaque département |
| Lundi 07:30 | Mise à jour Cockpit DG — vue hebdo | COCKPIT_DG.sheets Feuille 1 |
| 1er du mois 08:00 | Rapport mensuel complet | DG — Gmail + PowerPoint |

## Automatisation Gmail — Communications Liées au Système

### Rapports Automatiques par Rôle

**Responsables Départementaux — Email quotidien 19h00**
```
Objet : [BRASUS] Votre bilan du [Date] — [Département]

Bonjour [Prénom],

Votre taux d'exécution aujourd'hui : [X%]
Vos missions de demain seront disponibles dans votre espace à partir de 7h00.

[Si taux < 60%] : Un plan d'amélioration personnalisé a été préparé pour vous.

Bonne soirée,
BRASUS — {{NOM_ENTREPRISE}}
```

**Directeur Général — Email quotidien 19h30**
```
Objet : [BRASUS] Rapport de Direction — [Date]

Synthèse exécutive de la journée + alertes + plan d'actions.
Voir Cockpit DG pour les détails complets.
```

### Alertes Critiques (Temps Réel — toute la journée)

Déclenchement automatique si :
- Trésorerie < seuil d'alerte défini
- Taux d'exécution global < 40% à 15h00
- Aucune saisie d'un responsable après 14h00
- Incident opérationnel déclaré dans Onglet 2

## Google Calendar — Planification Système

### Événements Récurrents Créés par BRASUS

| Événement | Fréquence | Participants | Description |
|-----------|-----------|--------------|-------------|
| Clôture journalière | Quotidien 17h45 | Tous responsables | Rappel : saisir avant 18h00 |
| Revue performance dept | Hebdomadaire lundi 9h | Responsable + DG | 30 min — données BRASUS |
| Revue stratégique mensuelle | 1er lundi du mois | Équipe direction | 90 min — rapport mensuel |
| Alerte calendrier | Ad hoc | Responsable concerné | Si KPI critique détecté |

## Procédures d'Administration BRASUS

### Accès Administrateur Uniquement

Les opérations suivantes sont réservées à l'Administrateur BRASUS (DMA) :

- ✅ Créer un nouveau fichier départemental
- ✅ Configurer les permissions Google Workspace
- ✅ Modifier les missions générées (cas exceptionnel documenté)
- ✅ Déverrouiller une journée clôturée (procédure d'urgence)
- ✅ Ajouter / supprimer un responsable
- ✅ Modifier les seuils de performance
- ✅ Accéder aux logs BRASUS
- ✅ Modifier les templates de missions
- ✅ Configurer les automatisations Claude

### Procédure de Changement de Responsable

1. Révoquer l'accès de l'ancien responsable (immédiat)
2. Archiver son espace de travail (lecture seule préservée)
3. Créer l'accès du nouveau responsable
4. Former le nouveau responsable (guide rapide — 30 min)
5. Générer les missions d'intégration (premières 5 journées allégées)
6. Logger le changement dans BRASUS_CORE_CONFIG

## Sécurité & Conformité

### Règles de Sécurité Google Workspace

- Authentification Google obligatoire (pas de partage par lien)
- 2FA recommandée pour tous les comptes
- Session automatiquement fermée après 8h d'inactivité
- Aucun téléchargement des données départementales par les responsables
- Historique des accès conservé 24 mois

### Propriété des Données

- Les données de l'entreprise cliente = propriété de l'entreprise cliente
- Le moteur BRASUS, les agents, les procédures = propriété de DMA
- L'entreprise conserve l'accès à ses données même en fin de contrat
- Le moteur BRASUS est retiré en fin de contrat

## Variables de Configuration (depuis client-profile.md)

```yaml
google_workspace:
  domaine_client: "{{DOMAINE_GOOGLE}}"
  email_admin_brasus: "{{EMAIL_ADMIN_DMA}}"
  
  departements_actifs:
    - id: "commercial"
      email_responsable: "{{EMAIL_COMMERCIAL}}"
      sheet_id: "{{SHEET_ID_COMMERCIAL}}"
    - id: "marketing"
      email_responsable: "{{EMAIL_MARKETING}}"
      sheet_id: "{{SHEET_ID_MARKETING}}"
    - id: "rh"
      email_responsable: "{{EMAIL_RH}}"
      sheet_id: "{{SHEET_ID_RH}}"
    - id: "finance"
      email_responsable: "{{EMAIL_FINANCE}}"
      sheet_id: "{{SHEET_ID_FINANCE}}"
    - id: "operations"
      email_responsable: "{{EMAIL_OPERATIONS}}"
      sheet_id: "{{SHEET_ID_OPERATIONS}}"
  
  cockpit_dg:
    email_dg: "{{EMAIL_DG}}"
    sheet_id: "{{SHEET_ID_COCKPIT}}"
  
  horaires:
    generation_missions: "07:00"
    cloture_journaliere: "18:00"
    envoi_rapports: "18:30"
    fuseau_horaire: "Africa/Abidjan"
  
  seuils_performance:
    excellent: 90
    bon: 75
    moyen: 60
    insuffisant: 40
    critique: 0
  
  alertes:
    seuil_tresorerie_fcfa: "{{SEUIL_TRESORERIE}}"
    taux_execution_alerte: 40
    heure_alerte_non_saisie: "14:00"
```

---

SKILL généré par Digital Mastery Agency (DMA) — Abidjan, Côte d'Ivoire
BRASUS v2.0 — Système d'Exploitation d'Entreprise IA
Ce document est la propriété de DMA. Toute reproduction sans autorisation est interdite.
