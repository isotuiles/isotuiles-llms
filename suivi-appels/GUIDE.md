# Suivi des appels clients Isotuiles

Ce dossier contient tout ce qu'il faut pour ne plus perdre un appel client, sans abonnement payant.

Le cœur du système est le Google Sheets « Suivi appels clients Isotuiles » (dans votre Drive). Il contient l'onglet Appels (un dossier par client, avec motif, coordonnées, étape et date de rappel), un tableau de bord, l'onglet Listes (listes déroulantes et réglages) et un mode d'emploi.

Le script `apps-script/Code.gs` s'installe dans ce Sheets et fonctionne gratuitement sur votre compte Google. Il envoie l'email récapitulatif chaque matin, crée et déplace les rendez-vous « Rappeler … » dans Google Agenda, complète les fiches quand vous tapez un numéro et regroupe les appels d'un même numéro dans un seul dossier. Il fournit aussi le formulaire mobile `apps-script/Formulaire.html` et l'adresse web qui reçoit les appels manqués et alimente l'ERP.

Le fichier `isoteam/suivi-appels.html` est l'interface visuelle à intégrer dans Isotéam : un tableau par étape (Nouveau, À rappeler, Devis à faire…), les retards en rouge, un bouton pour appeler, les boutons « Rappelé » et « Demain », et le glisser-déposer pour changer d'étape. Ouvert tel quel, il affiche des données de démonstration.

## Installation

Les étapes détaillées sont dans l'onglet « Mode d'emploi » du Sheets. En résumé : collez le script dans Extensions > Apps Script, cliquez sur Isotuiles > Installer les automatisations, choisissez votre clé d'accès dans l'onglet Listes, déployez l'application web, puis renseignez `URL_API` et `CLE` en haut de `suivi-appels.html`.

## SMS automatique et appels manqués (Android, gratuit)

Avec l'application MacroDroid, créez une macro avec le déclencheur « Appel manqué » et deux actions. La première est une requête HTTP GET vers `URL_EXEC?action=manque&cle=VOTRE_CLE&tel={call_number}`, qui crée le dossier dans le Sheets et le rendez-vous de rappel. La seconde envoie un SMS au `{call_number}`, depuis votre forfait, avec par exemple ce texte :

> Bonjour, vous avez appelé Isotuiles et nous n'avons pas pu décrocher. Nous vous rappelons au plus vite. Pour gagner du temps, répondez à ce SMS avec votre nom, l'objet de votre appel (devis, commande, question) et la ville du chantier. Merci, l'équipe Isotuiles.

Ajoutez une contrainte « numéro absent des contacts » si vous ne voulez pas l'envoyer à vos proches. Sur iPhone, iOS ne permet pas cette automatisation : utilisez le formulaire mobile ajouté à l'écran d'accueil.

## Annonce de messagerie conseillée

> Bonjour, vous êtes bien chez Isotuiles. Nous sommes sur un chantier ou déjà en ligne. Laissez votre nom, votre numéro, l'objet de votre appel (devis, suivi de commande ou question) et la ville du chantier : nous vous rappelons dans la journée. Vous pouvez aussi demander un devis sur isotuiles.fr. Merci et à bientôt.

## Plus tard

Si vous voulez un vrai menu vocal (« tapez 1 pour un devis ») et la transcription automatique des messages, le script est prêt à recevoir les appels d'un numéro virtuel Twilio (quelques euros par mois) sans changer le Sheets ni l'ERP.
