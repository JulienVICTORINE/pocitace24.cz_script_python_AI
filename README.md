# 🧠 Počítače24.cz - AI Product Description Generator (Shoptet / E-commerce)

## 📌 Description

Ce projet permet de générer des descriptions longues de produits informatiques reconditionnés à partir de données structurées (CSV).

Il est conçu pour des plateformes e-commerce comme **Shoptet**, avec un objectif clair : produire des contenus HTML cohérents, professionnels et exploitables directement en production.

---

## 🎯 Objectif

Automatiser la rédaction de fiches produits en respectant :

* une structure rédactionnelle stricte
* des contraintes métier précises
* un ton professionnel adapté au e-commerce

Le projet garantit des descriptions homogènes, fiables et adaptées à des volumes importants de produits.

---

## ⚙️ Fonctionnalités

* Génération automatique de descriptions produits longues
* Utilisation de l’IA pour produire du contenu structuré
* Lecture et traitement de fichiers `CSV` (products)
* Adaptation selon la catégorie (ordinateur, laptop, écran…)
* Génération de FAQ dynamique
* Nettoyage et formatage HTML automatique

---

## Logique du projet

Le script :

1. Lit les données depuis `products.csv`
2. Analyse les caractéristiques techniques
3. Construit un prompt structuré
4. Génère une description via IA
5. Formate le résultat en **HTML propre**
6. Enregistre la sortie dans `produkt-ntb.html`

---

## 🏗️ Structure des descriptions générées

Chaque description suit une structure stricte :

1. Introduction
2. Présentation de la marque (ex : Lenovo)
3. Présentation de la gamme (ex : ThinkCentre)
4. Description de la configuration matérielle
5. Cas d’usage (Convient pour…)
6. Limitations (Ne convient pas pour…)
7. Produit prêt à l’emploi (Windows + garantie)

---

### 🛠️ Technologies utilisées

* Python
* OpenAI API
* Gemini API
* Regex / parsing HTML

---

## 📂 Structure du projet

```bash
.
├── merge_descriptions.py      # script principal de génération
├── product.csv                # données source
├── out_html/                  # fichiers HTML individuels
│   ├── product_1.html
│   ├── product_2.html
├── produkt-ntb.html           # fichier final consolidé
└── template_product.html      # modèle HTML + CSS 
```

---

## Composition du projet

### 📁 Output

Le projet génère des fichiers HTML dans le dossier suivant :

- `out_html/` : contient les descriptions produits générées individuellement
- `produkt-ntb.html` : fichier final consolidé

Le dossier `out_html` peut être utilisé comme étape intermédiaire pour :
- vérifier les descriptions individuellement
- faciliter le debug
- traiter les fichiers avant fusion


### 📝 Template HTML

Le fichier `template_product.html` sert de modèle pour générer les descriptions produits.

- Contient la structure de base HTML avec balises `<p>` et `<strong>` pour les sections obligatoires
- Permet de standardiser les descriptions
- Le script `merge_descriptions.py` remplace dynamiquement les parties du template par les informations spécifiques à chaque produit
- Facilite la mise à jour du formatage sans toucher au code Python


### 🖥️ merge_descriptions.py

Ce script Python est le moteur de génération des descriptions produits.

- Lit les données depuis `product.csv`
- Transforme les informations techniques en prompts structurés
- Utilise l’IA pour générer des descriptions longues, professionnelles et conformes aux règles métier
- Remplit le template `template_product.html` avec le contenu généré
- Gère l’export final dans `produkt-ntb.html`
- Optionnellement, crée des fichiers intermédiaires dans `out_html/` pour vérification ou debug


### 📄 produkt-ntb.html

Fichier HTML final qui contient toutes les descriptions produits générées.

- Consolidation de toutes les configurations dans un seul fichier
- Structure HTML standardisée avec `<p>` et `<strong>`
- Prêt à être intégré directement dans la boutique en ligne
- Respecte toutes les règles métier et la structure imposée par le projet

---

## 🚀 Installation

1. Cloner le repository :

```bash
[git clone https://github.com/ton-username/ton-repo.git](https://github.com/JulienVICTORINE/pocitace24.cz_script_python_AI.git)
cd pocitace24.cz_script_python_AI
```

2. Installer les dépendances :

```bash
pip install -r
```

3. Configurer les variables d’environnement :

```env
OPENAI_API_KEY=your_api_key
GEMINI_API_KEY=your_api_key
```

---

## 💡 Cas d’usage

* Boutiques e-commerce (Shoptet, Prestashop, Shopify…)
* Automatisation de contenu produit
* Enrichissement SEO à grande échelle
* Migration ou refonte de catalogue

---

## ⚠️ Contraintes importantes

* Les données proviennent exclusivement du fichier CSV
* Aucune génération hors cadre métier autorisé
* Respect strict du format HTML imposé
* Une seule sortie : `produkt-ntb.html`

---

## Auteur

Projet développé dans le cadre de mon stage chez la société Počítače24.cz à Prague
