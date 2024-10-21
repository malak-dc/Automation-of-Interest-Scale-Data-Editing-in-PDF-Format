

# 📄 Automatisation de l'Édition des Données d'Échelles d'Intérêts en Format PDF

## 🔍 Description

Ce projet vise à automatiser l'édition des échelles d'intérêts en format PDF pour le **Crédit Agricole du Maroc**. L'application recueille des données depuis une base **MySQL**, les traite en utilisant **Python** pour générer des fichiers Excel et PDF, et fournit une interface utilisateur développée avec **Java/Spring Boot** pour une gestion efficace des requêtes.

## 🎯 Fonctionnalités

- **Automatisation des Rapports** : Extraction des données MySQL, génération de fichiers Excel et conversion en PDF.
- **Interface Utilisateur** : Utilisation de **Java/Spring Boot** pour créer une interface web intuitive pour les utilisateurs.
- **Sécurité et Conformité** : Vérification rigoureuse des données pour assurer la validité et la conformité des rapports.
- **Technologies Utilisées** :
  - **MySQL** pour la gestion des données.
  - **Python** pour la génération des rapports Excel et leur conversion en PDF.
  - **Java/Spring Boot** pour l'interface backend.

## 📂 Structure du Projet

- **Backend** : Développé avec **Java/Spring Boot** pour gérer les requêtes utilisateur.
- **Scripts Python** : Utilisation des bibliothèques `openpyxl` et `win32com.client` pour manipuler les fichiers Excel et les convertir en PDF.
- **Base de Données** : Gestion et stockage des données dans **MySQL**.

## 🛠️ Installation

1. **Cloner le dépôt** :
   ```bash
   git clone https://github.com/votre-utilisateur/votre-projet.git
   cd votre-projet
   ```

2. **Configurer la base de données MySQL** :
   - Créez la base de données MySQL et importez les tables nécessaires.
   - Mettez à jour le fichier `application.properties` avec vos informations MySQL.

3. **Exécuter le backend** :
   ```bash
   mvn spring-boot:run
   ```

4. **Lancer les scripts Python** :
   - Installez les dépendances nécessaires :
     ```bash
     pip install openpyxl pywin32
     ```

## 💻 Utilisation

1. Ouvrez l'interface web : `http://localhost:8080`.
2. Recherchez un compte bancaire avec le numéro de compte.
3. Sélectionnez une période et générez le rapport PDF.

## 🖼️ Diagrammes

- **Diagramme d’Architecture** : Vue globale du système et des interactions entre les composants.
- **Diagramme UML** : Montre les relations entre les classes du projet.

## 📝 Résultats

- **Fichier Excel** : Les données sont extraites et formatées dans un fichier Excel.
- **Rapport PDF** : Conversion du fichier Excel en PDF pour une distribution simple et professionnelle.

## ✅ Conclusion

Ce projet a permis de **simplifier** et **accélérer** la génération des rapports PDF, tout en garantissant la précision, la sécurité, et la conformité des données. Grâce à l'automatisation et à l'intégration des différentes technologies, l'expérience utilisateur est grandement améliorée.

