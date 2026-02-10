#!/usr/bin/env python3
"""
Script to access and query the LISTE ARTICLE.xlsx file.
This script provides functionality to read, display, and search through the article list.
"""

import pandas as pd
import sys
import argparse
from pathlib import Path


class ArticleListReader:
    """Class to read and interact with the article list Excel file."""
    
    def __init__(self, file_path='LISTE ARTICLE.xlsx'):
        """
        Initialize the reader with the Excel file.
        
        Args:
            file_path: Path to the Excel file
        """
        self.file_path = Path(file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"Excel file not found: {file_path}")
        
        # Read the Excel file, skipping the header rows
        # The actual data starts at row 6 (index 5), with the column headers at row 5
        self.df_raw = pd.read_excel(file_path, header=None)
        
        # Extract the actual data with proper headers
        # Looking at the structure, row 5 has the column names
        self.df = self._parse_data()
        
    def _parse_data(self):
        """Parse the Excel data to extract the article information."""
        # The header row is at index 5 (6th row)
        # Read from row 6 onwards (data starts at row 7)
        df = pd.read_excel(self.file_path, skiprows=6)
        
        # Clean up column names - use only the meaningful ones
        # Based on inspection, the important columns are:
        # - Column 0: Référence (Reference)
        # - Column 3: Désignation (Description)
        # - Column 10: Famille (Family/Category)
        # - Column 12: Unité vente (Sales Unit)
        # - Column 17: Prix (Price)
        
        # Rename the first column to 'Référence'
        if len(df.columns) > 0:
            df = df.rename(columns={df.columns[0]: 'Référence'})
        
        # Create a cleaner dataframe with just the key columns
        clean_df = pd.DataFrame()
        
        if 'Référence' in df.columns:
            clean_df['Référence'] = df['Référence']
        
        # Try to find and rename other columns based on the data
        if len(df.columns) > 3:
            clean_df['Désignation'] = df.iloc[:, 3]
        if len(df.columns) > 10:
            clean_df['Famille'] = df.iloc[:, 10]
        if len(df.columns) > 12:
            clean_df['Unité_vente'] = df.iloc[:, 12]
        if len(df.columns) > 17:
            clean_df['Prix'] = df.iloc[:, 17]
        
        # Remove rows with all NaN values
        clean_df = clean_df.dropna(how='all')
        
        return clean_df
    
    def get_all_articles(self):
        """Return all articles as a DataFrame."""
        return self.df
    
    def search_by_reference(self, reference):
        """
        Search for an article by reference number.
        
        Args:
            reference: The reference number to search for
            
        Returns:
            DataFrame with matching articles
        """
        if 'Référence' not in self.df.columns:
            print("Warning: 'Référence' column not found")
            return pd.DataFrame()
        
        return self.df[self.df['Référence'].astype(str).str.contains(str(reference), case=False, na=False)]
    
    def search_by_description(self, keyword):
        """
        Search for articles by description keyword.
        
        Args:
            keyword: Keyword to search in description
            
        Returns:
            DataFrame with matching articles
        """
        if 'Désignation' not in self.df.columns:
            print("Warning: 'Désignation' column not found")
            return pd.DataFrame()
        
        return self.df[self.df['Désignation'].astype(str).str.contains(keyword, case=False, na=False)]
    
    def get_by_family(self, family):
        """
        Get all articles from a specific family.
        
        Args:
            family: The family/category name
            
        Returns:
            DataFrame with articles from the specified family
        """
        if 'Famille' not in self.df.columns:
            print("Warning: 'Famille' column not found")
            return pd.DataFrame()
        
        return self.df[self.df['Famille'].astype(str).str.contains(family, case=False, na=False)]
    
    def get_statistics(self):
        """Get basic statistics about the article list."""
        stats = {
            'total_articles': len(self.df),
            'columns': list(self.df.columns),
            'families': self.df['Famille'].nunique() if 'Famille' in self.df.columns else 'N/A',
        }
        
        if 'Prix' in self.df.columns:
            # Convert Prix to numeric, handling any non-numeric values
            prix_numeric = pd.to_numeric(self.df['Prix'], errors='coerce')
            stats['min_price'] = prix_numeric.min()
            stats['max_price'] = prix_numeric.max()
            stats['avg_price'] = prix_numeric.mean()
        
        return stats
    
    def display_summary(self, n=10):
        """
        Display a summary of the first n articles.
        
        Args:
            n: Number of articles to display
        """
        print(f"\n{'='*80}")
        print(f"LISTE ARTICLE - Summary")
        print(f"{'='*80}")
        
        stats = self.get_statistics()
        print(f"\nTotal articles: {stats['total_articles']}")
        print(f"Columns: {', '.join(stats['columns'])}")
        
        if isinstance(stats.get('families'), int):
            print(f"Number of families: {stats['families']}")
        
        if 'min_price' in stats:
            print(f"\nPrice range: {stats['min_price']:.0f} - {stats['max_price']:.0f}")
            print(f"Average price: {stats['avg_price']:.2f}")
        
        print(f"\nFirst {n} articles:")
        print(self.df.head(n).to_string(index=False))
        print(f"\n{'='*80}\n")


def main():
    """Main function to run the script from command line."""
    parser = argparse.ArgumentParser(description='Access and query the LISTE ARTICLE Excel file')
    parser.add_argument('--file', default='LISTE ARTICLE.xlsx', help='Path to the Excel file')
    parser.add_argument('--summary', action='store_true', help='Display summary of the file')
    parser.add_argument('--search-ref', help='Search by reference number')
    parser.add_argument('--search-desc', help='Search by description keyword')
    parser.add_argument('--family', help='Filter by family/category')
    parser.add_argument('--all', action='store_true', help='Display all articles')
    parser.add_argument('--stats', action='store_true', help='Display statistics')
    parser.add_argument('--limit', type=int, default=50, help='Limit number of results to display')
    
    args = parser.parse_args()
    
    try:
        # Create reader instance
        reader = ArticleListReader(args.file)
        
        # If no specific action is specified, show summary
        if not any([args.summary, args.search_ref, args.search_desc, args.family, args.all, args.stats]):
            reader.display_summary()
            return
        
        # Handle different actions
        if args.summary:
            reader.display_summary(args.limit)
        
        if args.stats:
            stats = reader.get_statistics()
            print("\nStatistics:")
            for key, value in stats.items():
                print(f"  {key}: {value}")
        
        if args.search_ref:
            results = reader.search_by_reference(args.search_ref)
            print(f"\nSearch results for reference '{args.search_ref}':")
            print(results.head(args.limit).to_string(index=False))
            print(f"\nFound {len(results)} matching articles")
        
        if args.search_desc:
            results = reader.search_by_description(args.search_desc)
            print(f"\nSearch results for description '{args.search_desc}':")
            print(results.head(args.limit).to_string(index=False))
            print(f"\nFound {len(results)} matching articles")
        
        if args.family:
            results = reader.get_by_family(args.family)
            print(f"\nArticles in family '{args.family}':")
            print(results.head(args.limit).to_string(index=False))
            print(f"\nFound {len(results)} articles in this family")
        
        if args.all:
            print("\nAll articles:")
            print(reader.get_all_articles().head(args.limit).to_string(index=False))
            print(f"\nShowing {min(args.limit, len(reader.df))} of {len(reader.df)} articles")
    
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
