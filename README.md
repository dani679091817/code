# Accès au Fichier Excel - LISTE ARTICLE

Ce dépôt contient un outil Python pour accéder et interroger le fichier Excel "LISTE ARTICLE.xlsx".

## Description

Le fichier `LISTE ARTICLE.xlsx` contient une liste d'articles avec les informations suivantes :
- **Référence** : Numéro de référence de l'article
- **Désignation** : Description de l'article
- **Famille** : Catégorie/famille de l'article
- **Unité vente** : Unité de vente
- **Prix** : Prix de l'article

## Installation

### Prérequis
- Python 3.7 ou supérieur
- pip (gestionnaire de paquets Python)

### Installation des dépendances

```bash
pip install -r requirements.txt
```

## Utilisation

### 1. Afficher un résumé du fichier

Pour afficher un résumé avec les 10 premiers articles :

```bash
python access_excel.py
```

ou

```bash
python access_excel.py --summary
```

### 2. Afficher les statistiques

Pour afficher les statistiques globales :

```bash
python access_excel.py --stats
```

### 3. Rechercher par référence

Pour rechercher un article par son numéro de référence :

```bash
python access_excel.py --search-ref ACB00001
```

### 4. Rechercher par description

Pour rechercher des articles par mot-clé dans la description :

```bash
python access_excel.py --search-desc "ASSIETTE"
```

### 5. Filtrer par famille

Pour afficher tous les articles d'une famille spécifique :

```bash
python access_excel.py --family ACB
```

### 6. Afficher tous les articles

Pour afficher tous les articles (avec une limite) :

```bash
python access_excel.py --all --limit 100
```

### Options disponibles

- `--file` : Chemin vers le fichier Excel (par défaut : "LISTE ARTICLE.xlsx")
- `--summary` : Afficher un résumé du fichier
- `--search-ref` : Rechercher par numéro de référence
- `--search-desc` : Rechercher par mot-clé dans la description
- `--family` : Filtrer par famille/catégorie
- `--all` : Afficher tous les articles
- `--stats` : Afficher les statistiques
- `--limit` : Limiter le nombre de résultats affichés (par défaut : 50)

## Utilisation dans du code Python

Vous pouvez également utiliser la classe `ArticleListReader` dans votre propre code Python :

```python
from access_excel import ArticleListReader

# Créer une instance du lecteur
reader = ArticleListReader('LISTE ARTICLE.xlsx')

# Obtenir tous les articles
articles = reader.get_all_articles()
print(articles)

# Rechercher par référence
results = reader.search_by_reference('ACB00001')
print(results)

# Rechercher par description
results = reader.search_by_description('ASSIETTE')
print(results)

# Obtenir les articles d'une famille
results = reader.get_by_family('ACB')
print(results)

# Obtenir les statistiques
stats = reader.get_statistics()
print(stats)
```

## Structure du fichier Excel

Le fichier Excel contient environ 8 000+ articles avec les colonnes suivantes :
- Référence
- Désignation (description du produit)
- Famille (catégorie)
- Unité de vente
- Prix

## Exemples de résultats

```
================================================================================
LISTE ARTICLE - Summary
================================================================================

Total articles: 8132
Columns: Référence, Désignation, Famille, Unité_vente, Prix

First 10 articles:
 Référence                             Désignation Famille Unité_vente    Prix
  ACB00001                      ADAPTATEUR MARKEN     ACB      pieces     300
  ACB00002                  ALLUME GAZ + RECHARGE     ACB      pieces    1500
  ACB00003           ASSIETE CASSABLE DESSINI L*3     ACB      pieces   18000
...
================================================================================
```

## Licence

Ce projet est destiné à un usage interne pour gérer la liste d'articles.
