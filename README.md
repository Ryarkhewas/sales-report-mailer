# rapport-auto

**Transforme un export CSV de ventes en rapport Excel prêt à présenter, et l'envoie par email automatiquement.**

Beaucoup de petites entreprises passent une demi-journée par mois à copier-coller leurs exports dans Excel, faire des totaux et des graphiques, puis envoyer le tout par mail. Ce script fait les trois en une commande, et peut tourner tout seul chaque semaine ou chaque mois.

## Ce que fait le script

1. **Lit un export CSV** (séparateur `;` ou `,`, décimales à la française acceptées) et ignore proprement les lignes invalides.
2. **Calcule les indicateurs clés** : chiffre d'affaires, nombre de ventes, clients distincts, panier moyen, évolution mensuelle, répartition par produit, top clients.
3. **Génère un fichier Excel mis en forme** :
   - une feuille *Synthèse* avec les indicateurs, les tableaux et deux graphiques ;
   - une feuille *Détail des ventes* avec filtres et en-tête figé ;
   - une mise en page prête à imprimer (A4, ajusté en largeur).
4. **Envoie le rapport par email** (optionnel) avec un résumé dans le corps du message et l'Excel en pièce jointe.

## Installation

```bash
git clone https://github.com/Ryarkhewas/rapport-auto.git
cd rapport-auto
python -m venv .venv
source .venv/bin/activate      # Windows : .venv\Scripts\activate
pip install -r requirements.txt
```

## Utilisation

```bash
# Générer le rapport à partir des données d'exemple
python rapport.py data/ventes_exemple.csv

# Choisir le fichier de sortie et le titre
python rapport.py data/ventes_exemple.csv -o rapports/septembre.xlsx -t "Ventes de septembre"

# Préparer l'email sans l'envoyer (pour vérifier)
python rapport.py data/ventes_exemple.csv --envoyer --dry-run

# Générer et envoyer
python rapport.py data/ventes_exemple.csv --envoyer
```

## Format du CSV attendu

| date       | client          | produit               | quantite | prix_unitaire |
|------------|-----------------|-----------------------|----------|---------------|
| 01/04/2026 | Café des Arts   | Support à distance    | 3        | 45,00         |

Les colonnes supplémentaires sont ignorées. Les dates au format `JJ/MM/AAAA` ou `AAAA-MM-JJ` sont acceptées.

## Configuration de l'envoi d'email

Copiez `.env.example` en `.env` et renseignez vos paramètres SMTP :

```bash
cp .env.example .env
```

Avec Gmail, utilisez un **mot de passe d'application** (Compte Google → Sécurité → Validation en deux étapes → Mots de passe des applications), jamais votre mot de passe principal. Le fichier `.env` est exclu de Git.

## Automatiser l'exécution

**Linux / macOS (cron)** — chaque 1er du mois à 8 h :

```cron
0 8 1 * * cd /chemin/vers/rapport-auto && .venv/bin/python rapport.py /chemin/export.csv --envoyer
```

**Windows** : Planificateur de tâches → Créer une tâche de base → action « Démarrer un programme » avec `python.exe` et les arguments `rapport.py C:\chemin\export.csv --envoyer`.

## Tests

```bash
python -m pytest -q
```

Les tests couvrent le chargement (formats, lignes invalides, colonnes manquantes), les calculs, la génération Excel, la construction de l'email et la ligne de commande.

## Sécurité

- Les identifiants SMTP sont lus depuis l'environnement, jamais écrits dans le code.
- Connexion SMTP chiffrée (STARTTLS sur le port 587, SSL sur le 465).
- Les données d'exemple sont fictives.

## Adapter à votre besoin

Le script est volontairement simple à modifier : autres colonnes, autres indicateurs, export PDF, envoi vers Teams ou Slack, lecture directe depuis une base de données ou une API… [Contactez-moi](https://github.com/Ryarkhewas) pour une version sur mesure.

## Licence

MIT
