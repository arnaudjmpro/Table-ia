# Installation de la prise de photo TableIA

Cette version conserve la galerie et la dictée existantes. Dans Excel sur Mac,
le bouton **Prendre une photo** ouvre la caméra dans le navigateur, puis la photo
revient automatiquement dans le panneau TableIA.

## 1. Créer la table temporaire dans Supabase

Dans **Supabase > SQL Editor**, exécuter le fichier :

`supabase/migrations/20260927090000_tableia_photo_sessions.sql`

## 2. Déployer la nouvelle Edge Function

Créer une fonction nommée exactement :

`tableia-photo-bridge`

Puis utiliser le contenu du fichier :

`supabase/functions/tableia-photo-bridge/index.ts`

Les variables `SUPABASE_URL` et `SUPABASE_SERVICE_ROLE_KEY` sont fournies
automatiquement par Supabase aux Edge Functions. La clé de service ne doit
jamais être copiée dans les fichiers publics du complément.

## 3. Mettre à jour le dépôt TableIA

Remplacer les fichiers publics par ceux du paquet :

- `app.js`
- `dialog-camera.html`
- `index.html`
- `style.css`
- `manifest.xml`

Ajouter également le dossier `supabase` au dépôt afin de conserver le code
serveur et la migration.

## 4. Tester

1. Ouvrir TableIA dans Excel.
2. Appuyer sur **Prendre une photo**.
3. Autoriser la caméra dans le navigateur.
4. Prendre la photo.
5. Revenir dans Excel : l'aperçu doit apparaître automatiquement.

Les photos du relais expirent après dix minutes et sont supprimées dès leur
première récupération par Excel.
