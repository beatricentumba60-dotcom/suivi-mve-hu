# Suivi MVE – publication automatique

Ce dépôt contient :
- la branche **main** : le programme de publication et la tâche planifiée (.github/workflows/publier.yml) ;
- la branche **gh-pages** : le site publié (applications Suivi MVE et Qualité MVE), régénéré à chaque publication.

Toutes les 30 minutes, GitHub lit DHIS2, chiffre les données avec le mot de passe d'équipe et met le site à jour.
Les secrets (jeton DHIS2, mot de passe d'équipe) sont dans Settings › Secrets and variables › Actions.
Publier à la demande : onglet Actions › Publication Suivi MVE › Run workflow.
Aucune donnée en clair n'est déposée ici.
