# Guide d'utilisation de l'application Prétraitement de données brutes via Docker

Ce guide explique comment utiliser l'application de prétraitement de données brutes sans rien installer d'autre que Docker. Vous n'avez pas besoin de Python ni de dépendances spécifiques.

---

## 1. Installer Docker Desktop

- Rendez-vous sur : https://www.docker.com/products/docker-desktop/
- Téléchargez et installez Docker Desktop pour Windows ou Mac.
- Lancez Docker Desktop et attendez que l'icône devienne verte (Docker doit être "running").

## 2. Récupérer l'image Docker

### **Option A : Depuis un fichier fourni**
1. Placez le fichier `pretraitement-donneebrute.tar` reçu sur votre ordinateur (par exemple sur le Bureau ou dans Documents).
2. **Ne décompressez pas ce fichier !** Docker lit directement ce format.
3. Ouvrez un terminal (PowerShell ou Invite de commandes).
4. Chargez l'image Docker :
   - **Si vous êtes dans le dossier où se trouve le fichier :**
     ```bash
     docker load -i pretraitement-donneebrute.tar
     ```
   - **Si le fichier est ailleurs, indiquez le chemin complet :**
     ```bash
     docker load -i "C:\Users\VotreNom\Documents\pretraitement-donneebrute.tar"
     ```
     *(Remplacez `VotreNom` par votre nom d'utilisateur Windows. Les guillemets sont utiles si le chemin contient des espaces.)*
   - Vous pouvez lancer cette commande depuis n'importe quel dossier, il suffit d'indiquer le chemin complet du fichier `.tar` si besoin.

### **Option B : Depuis Docker Hub**
1. Ouvrez un terminal.
2. Téléchargez l'image (remplacez `tonpseudo` par le nom fourni) :
   ```bash
   docker pull tonpseudo/pretraitement-donneebrute:latest
   ```

## 3. Lancer l'application

Dans le terminal, tapez :
```bash
docker run -p 8501:8501 pretraitement-donneebrute
```
- Si vous avez téléchargé depuis Docker Hub, remplacez `pretraitement-donneebrute` par `tonpseudo/pretraitement-donneebrute:latest`.

## 4. Accéder à l'application

- Ouvrez votre navigateur internet.
- Rendez-vous sur : [http://localhost:8501](http://localhost:8501)

## 5. Arrêter l'application
- Dans le terminal où l'application tourne, faites `Ctrl+C`.
- Ou, dans un autre terminal, listez les conteneurs :
  ```bash
  docker ps
  ```
  puis arrêtez le conteneur avec :
  ```bash
  docker stop <container_id>
  ```

---

**Remarques :**
- Les fichiers uploadés et générés restent dans le conteneur (ils ne sont pas sauvegardés sur votre PC par défaut).
- Si vous avez un problème, vérifiez que Docker Desktop est bien lancé.
- Pour relancer l'application, refaites simplement l'étape 3.
- **Il n'est pas nécessaire de décompresser le fichier .tar, Docker s'en charge automatiquement.**
- **Important :** Le chargement (load) des fichiers Excel doit être fait une seule fois par traitement. Recharger plusieurs fois le même fichier dans le même processus peut provoquer des erreurs ou des incohérences.

---

## Exemple détaillé : Charger l'image Docker depuis un fichier .tar

1. **Copiez le fichier** `pretraitement-donneebrute.tar` dans un dossier facile à retrouver (par exemple, sur le Bureau ou dans Documents).
2. **Ouvrez un terminal** (PowerShell ou Invite de commandes).
3. **Allez dans le dossier** (optionnel) :
   ```bash
   cd C:\Users\VotreNom\Documents
   ```
4. **Ou indiquez le chemin complet** dans la commande :
   ```bash
   docker load -i "C:\Users\VotreNom\Documents\pretraitement-donneebrute.tar"
   ```
5. **Attendez la fin de l'import** (Docker affiche un message de succès).
6. **Lancez l'application** :
   ```bash
   docker run -p 8501:8501 pretraitement-donneebrute
   ```
7. **Ouvrez votre navigateur** et allez sur [http://localhost:8501](http://localhost:8501)

Pour toute question, contactez l'administrateur du projet. 