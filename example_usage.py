#!/usr/bin/env python3
"""
Exemple d'utilisation de la classe ArticleListReader.
Example script demonstrating how to use the ArticleListReader class programmatically.
"""

from access_excel import ArticleListReader


def main():
    """Démonstration de l'utilisation du module."""
    
    print("="*80)
    print("Exemple d'utilisation du module ArticleListReader")
    print("="*80)
    
    # Créer une instance du lecteur / Create reader instance
    reader = ArticleListReader('LISTE ARTICLE.xlsx')
    
    # 1. Afficher les statistiques / Display statistics
    print("\n1. Statistiques de la liste d'articles:")
    print("-" * 40)
    stats = reader.get_statistics()
    for key, value in stats.items():
        print(f"   {key}: {value}")
    
    # 2. Rechercher un article par référence / Search by reference
    print("\n2. Recherche par référence 'ACB00001':")
    print("-" * 40)
    result = reader.search_by_reference('ACB00001')
    if not result.empty:
        print(result.to_string(index=False))
    else:
        print("   Aucun article trouvé / No articles found")
    
    # 3. Rechercher par mot-clé dans la description / Search by description keyword
    print("\n3. Recherche par mot-clé 'BIC' dans la description:")
    print("-" * 40)
    results = reader.search_by_description('BIC')
    print(f"   Trouvé {len(results)} articles")
    if not results.empty:
        print("\n   Premiers résultats:")
        print(results.head(3).to_string(index=False))
    
    # 4. Obtenir les articles d'une famille / Get articles by family
    print("\n4. Articles de la famille 'ACB':")
    print("-" * 40)
    family_articles = reader.get_by_family('ACB')
    print(f"   Total: {len(family_articles)} articles dans cette famille")
    if not family_articles.empty:
        print("\n   Premiers articles:")
        print(family_articles.head(3).to_string(index=False))
    
    # 5. Obtenir tous les articles / Get all articles
    print("\n5. Aperçu de tous les articles:")
    print("-" * 40)
    all_articles = reader.get_all_articles()
    print(f"   Total: {len(all_articles)} articles")
    print("\n   Premiers articles:")
    print(all_articles.head(5).to_string(index=False))
    
    # 6. Exemple de filtrage personnalisé / Custom filtering example
    print("\n6. Exemple de filtrage: Articles avec prix > 10000:")
    print("-" * 40)
    import pandas as pd
    all_articles = reader.get_all_articles()
    if 'Prix' in all_articles.columns:
        # Convertir Prix en numérique / Convert Price to numeric
        all_articles['Prix_num'] = pd.to_numeric(all_articles['Prix'], errors='coerce')
        expensive_items = all_articles[all_articles['Prix_num'] > 10000]
        print(f"   Trouvé {len(expensive_items)} articles")
        if not expensive_items.empty:
            print("\n   Quelques exemples:")
            print(expensive_items[['Référence', 'Désignation', 'Prix']].head(5).to_string(index=False))
    
    print("\n" + "="*80)
    print("Fin de la démonstration / End of demonstration")
    print("="*80)


if __name__ == '__main__':
    main()
