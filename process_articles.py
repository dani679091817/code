#!/usr/bin/env python3
"""
Traitement de LISTE PRIX ARTICLE 2.xlsx
- Filtre les articles VIN... et WHI...
- Enrichit le tableau avec : Contenance (cl), Catégorie, Degré d'alcool, Nom complet
"""

import pandas as pd
import re


# ---------------------------------------------------------------------------
# Extraction de la contenance en centilitres
# ---------------------------------------------------------------------------

def extract_contenance(designation):
    """Extrait le volume en centilitres depuis la désignation."""
    text = str(designation).upper()

    # 1) pattern XXXcl (ex: 75CL, 100CL, 1.5CL)
    m = re.search(r'(\d+(?:[.,]\d+)?)\s*CL\b', text)
    if m:
        val = float(m.group(1).replace(',', '.'))
        # Correction des typos : 750CL -> 75 cl  (> 200 => probablement des ml)
        if val > 200:
            return round(val / 10, 1)
        # 1.5CL est probablement une typo pour 1.5L (150 cl)
        if val < 5 and '.' in m.group(1):
            return round(val * 100, 1)
        return val

    # 2) pattern XXXml (ex: 750ML, 700ML, 200ML)
    m = re.search(r'(\d+(?:[.,]\d+)?)\s*ML\b', text)
    if m:
        val = float(m.group(1).replace(',', '.'))
        return round(val / 10, 1)

    # 3) pattern X.XXL ou XL (ex: 0.75L, 1.5L, 1L)
    m = re.search(r'(\d+(?:[.,]\d+)?)\s*L\b', text)
    if m:
        val = float(m.group(1).replace(',', '.'))
        # On ignore les années et numéros (> 10 L = irréaliste)
        if val > 10:
            return None
        return round(val * 100, 1)

    return None


# ---------------------------------------------------------------------------
# Catégorisation selon la hiérarchie des boissons alcoolisées
# ---------------------------------------------------------------------------

def get_categorie(reference, designation):
    """Retourne le chemin de catégorie dans la hiérarchie des boissons alcoolisées."""
    ref = str(reference).upper()
    des = str(designation).upper()

    # --- Liqueurs / Cocktails / Apéritifs (avant tout) ---

    # Cocktails prêts-à-boire
    if 'COCKTAIL' in des or 'CANNETTE' in des:
        return '/Liqueurs'

    if any(kw in des for kw in ['APPERITIF', 'APERITIF', 'APERITIVO', 'APEROL']) and 'PASTIS' not in des:
        return '/Liqueurs'

    if any(kw in des for kw in ['VERMOUTH', 'VERMOUH', 'VERMUTI']):
        return '/Liqueurs'

    if any(kw in des for kw in ['MARTINI BIANCO', 'MARTINI ROSSO', 'MARTINI ROSATO',
                                 'MARTINI EXTRA', 'MARTINI FIERO']):
        return '/Liqueurs'

    if re.search(r'CREME\s+D[EUA]\b', des):
        return '/Liqueurs'

    if any(kw in des for kw in ['BAILEYS', 'LIMONCELLO', 'AMARETTO', 'SAMBUCA',
                                 'COINTREAU', 'GRAND MARNIER', 'TRIPLE SEC',
                                 'CURACAO', 'AMARO', 'BITTER ', 'CAMPARI',
                                 'DEMANDIS', 'MERRYS', 'SHERIDANS', 'YACHTING',
                                 'JARDINS SECRETS', 'PEPPERMINT', 'PIPPERMINT',
                                 'PEPERMINT']):
        return '/Liqueurs'

    if 'LIQUEUR' in des:
        return '/Liqueurs'

    # --- Spiritueux ---

    if 'TEQUILA' in des:
        return '/Spiritueux/Tequila'

    # PASTIS = spiritueux (45% alc.) à placer avant le check apéritif
    if 'PASTIS' in des:
        return '/Spiritueux/Brandy'

    if any(kw in des for kw in ['RHUM', ' RUM ', 'CAPTAIN MORGAN', 'OLD NICK']):
        return '/Spiritueux/Rhum'

    if 'VODKA' in des:
        return '/Spiritueux/Vodka'

    # GIN (exclure "GINETTO" = marque de boisson Hugo, "GINGER")
    if re.search(r'\bGIN\b', des) and 'GINETTO' not in des and 'GINGER' not in des:
        return '/Spiritueux/Gin'

    if any(kw in des for kw in ['COGNAC', 'COGNIAC', ' BRANDY', 'SUAU BRANDY']):
        return '/Spiritueux/Brandy'

    if any(kw in des for kw in ['WHISKY', 'WHISKEY', 'SCOTCH', 'BOURBON',
                                 'BALLANTINES', 'MONKEY SHOULDER', 'GLEN TURNER',
                                 'IMPERIAL BLUE', 'GLENFIDDICH', 'CHIVAS',
                                 'SIR EDWARDS', 'SIR EDWARD S']):
        return '/Spiritueux/Whisky'

    # --- Vins fortifiés ---
    if 'PORTO' in des:
        return '/Vins/Vin rouge'

    # --- Vins : détection indépendante de la référence ---

    # Vin mousseux (priorité sur blanc/rouge)
    mousseux_kw = ['MOUSSEUX', 'MOUSEUX', 'MOUSS ', 'MSX', 'PROSECCO',
                   'CHAMPAGNE', 'BRUT', 'CREMANT', 'CAVA', 'BULLES',
                   'FANT NIGHT', 'DE BLC MSX', 'BLC DE BLC',
                   'SOIR PARIS', 'MUSCADOR', 'VERRY ']
    if any(kw in des for kw in mousseux_kw):
        return '/Vins/Vin mousseux'

    # Rosé
    if re.search(r'\b(RSE|ROSE|ROSÉ|ROSATO|ROSA)\b', des):
        return '/Vins/Vin rosé'

    # Blanc
    blanc_kw = [r'\bBLC\b', r'\bBLANC\b', r'\bBIANCO\b', r'\bCHARDONNAY\b',
                r'\bSAUVIGNON\b', r'\bMUSCAT\b', r'\bMOSCATO\b', r'\bRIESLING\b',
                r'\bPINOT GRIGIO\b', r'\bCOLOMBARD\b', r'\bZINFANDEL\b',
                r'\bMONBAZILLAC\b', r'\bSAUTERNES?\b', r'\bSEMILLON\b',
                r'\bAIREN\b', r'\bGROS MANSENA?\b', r'\bMUSCADET\b',
                r'\bMOELLEUX\b', r'\bMOEL\b', r'\bMLX\b', r'\bMOELL\b',
                r'\bMOUELLEUX\b']
    if any(re.search(kw, des) for kw in blanc_kw):
        return '/Vins/Vin blanc'

    # Rouge
    rouge_kw = [r'\bRGE\b', r'\bROUGE\b', r'\bROSS[OA]?\b', r'\bTINTO\b',
                r'\bNEROAMARO\b', r'\bPRIMITIVO\b', r'\bMERLOT\b',
                r'\bCABERNET\b', r'\bCAB.SAUV\b', r'\bSHIRAZ\b',
                r'\bSYRAH\b', r'\bPINOT NOIR\b', r'\bMALBEC\b',
                r'\bSANGIOVESE\b', r'\bCASELAO\b', r'\bCARMENERE\b',
                r'\bTEMPRANILLO\b', r'\bCASTELAO\b']
    if any(re.search(kw, des) for kw in rouge_kw):
        return '/Vins/Vin rouge'

    # Par défaut selon la référence
    if ref.startswith('VIN'):
        return '/Vins/Vin rouge'

    if ref.startswith('WHI'):
        return '/Spiritueux/Whisky'

    return '/Boissons alcoolisées'


# ---------------------------------------------------------------------------
# Degré d'alcool estimé
# ---------------------------------------------------------------------------

def get_degre_alcool(categorie, designation):
    """Estime le degré d'alcool typique selon la catégorie et la désignation."""
    des = str(designation).upper()
    cat = str(categorie)

    # Produits sans alcool
    if re.search(r'\bSANS ALCOOL\b|\bS\.A\b|\bSA\b(?!UV)', des):
        return '0%'

    if '/Vins/Vin rouge' in cat:
        if 'PORTO' in des:
            return '20%'
        if any(kw in des for kw in ['SYRAH', 'SHIRAZ', 'PRIMITIVO', 'NEROAMARO']):
            return '14%'
        if any(kw in des for kw in ['MERLOT', 'CABERNET', 'CAB-SAUV', 'CAB SAUV', 'MALBEC']):
            return '13.5%'
        if any(kw in des for kw in ['MEDOC', 'POMEROL', 'EMILION', 'GRAVES', 'BDX', 'BORDEAUX']):
            return '13%'
        return '12.5%'

    elif '/Vins/Vin blanc' in cat:
        if any(kw in des for kw in ['MOSCATO', 'MUSCAT']):
            return '7%'
        if any(kw in des for kw in ['MONBAZILLAC', 'SAUTERNES', 'SAUTERNE']):
            return '13.5%'
        if any(kw in des for kw in ['MOELLEUX', 'MOEL', 'MLX', 'MOUELLEUX', 'MOELL']):
            return '11.5%'
        if 'CHARDONNAY' in des:
            return '13%'
        if 'SAUVIGNON' in des:
            return '12.5%'
        if 'PINOT GRIGIO' in des:
            return '12.5%'
        return '12%'

    elif '/Vins/Vin rosé' in cat:
        return '12%'

    elif '/Vins/Vin mousseux' in cat:
        if re.search(r'\bSANS ALCOOL\b|\bS\.A\b', des):
            return '0%'
        if any(kw in des for kw in ['MOSCATO', 'MUSCAT']):
            return '7%'
        if 'PROSECCO' in des:
            return '11%'
        return '12%'

    elif '/Spiritueux/Whisky' in cat:
        if 'BOURBON' in des:
            return '43%'
        if re.search(r'\bSINGL[E]? (MALT|BARREL|GRAIN)\b', des):
            return '43%'
        return '40%'

    elif '/Spiritueux/Rhum' in cat:
        return '40%'

    elif '/Spiritueux/Vodka' in cat:
        return '40%'

    elif '/Spiritueux/Gin' in cat:
        return '40%'

    elif '/Spiritueux/Tequila' in cat:
        return '38%'

    elif '/Spiritueux/Brandy' in cat:
        if any(kw in des for kw in ['COGNAC', 'COGNIAC']):
            return '40%'
        if 'PASTIS' in des:
            return '45%'
        return '36%'

    elif '/Liqueurs' in cat:
        if any(kw in des for kw in ['CREAM', 'CREMA', 'CREME', 'BAILEYS']):
            return '17%'
        if any(kw in des for kw in ['TRIPLE SEC', 'CURACAO', 'COINTREAU', 'GRAND MARNIER']):
            return '40%'
        if any(kw in des for kw in ['LIMONCELLO', 'SAMBUCA', 'AMARETTO']):
            return '28%'
        if any(kw in des for kw in ['APERITIF', 'APPERITIF', 'APEROL', 'CAMPARI',
                                     'APERITIVO', 'MARTINI', 'VERMOUTH', 'PASTIS']):
            return '15%'
        if any(kw in des for kw in ['BITTER', 'AMARO']):
            return '30%'
        return '20%'

    return 'N/A'


# ---------------------------------------------------------------------------
# Expansion des abréviations → Nom complet
# ---------------------------------------------------------------------------

# Remplacement de séquences multi-mots (ordre décroissant de longueur)
PHRASE_MAP = [
    ('CAB-SAUV',          'Cabernet Sauvignon'),
    ('CAB SAUV',          'Cabernet Sauvignon'),
    ('B.S.W',             'Blended Scotch Whisky'),
    ('C-D-R',             'Côtes du Rhône'),
    ('CTES DU RHONE',     'Côtes du Rhône'),
    ('HT MEDOC',          'Haut-Médoc'),
    ('HT-MEDOC',          'Haut-Médoc'),
    ('BDX SUP',           'Bordeaux Supérieur'),
    ('LUSSAC-ST-EMILION', 'Lussac Saint-Émilion'),
    ('ST EMILION',        'Saint-Émilion'),
    ('ST-EMILION',        'Saint-Émilion'),
    ('GD CRU',            'Grand Cru'),
    ('CTES DE BERGERAC',  'Côtes de Bergerac'),
    ('CTES DE BOURG',     'Côtes de Bourg'),
    ('CTES DU THAU',      'Côtes du Thau'),
    ('VIN DE FRANCE',     'Vin de France'),
    ('SANG ',             'Sangiovese '),
    (r'CAISSE\*(\d+)X(\d+)CL', lambda m: f'Caisse de {m.group(1)} x {m.group(2)}cl'),
    (r'CAISSE\*(\d+)',     lambda m: f'Caisse de {m.group(1)}'),
    (r'LOT\*\s*(\d+)',    lambda m: f'Lot de {m.group(1)}'),
]

# Remplacement mot à mot
WORD_MAP = {
    # Couleur / type de vin
    'RGE':       'Rouge',
    'BLC':       'Blanc',
    'RSE':       'Rosé',
    'MSX':       'Mousseux',
    'MOUSS':     'Mousseux',
    'MOUSSEUX':  'Mousseux',
    'MOUSEUX':   'Mousseux',
    # Texture
    'MOEL':      'Moelleux',
    'MOELL':     'Moelleux',
    'MLX':       'Moelleux',
    'MOELLEUX':  'Moelleux',
    'MOUELLEUX': 'Moelleux',
    'MOUELL':    'Moelleux',
    'MLLX':      'Moelleux',
    'MLLUX':     'Moelleux',
    'MLLEUX':    'Moelleux',
    'SEC':       'Sec',
    # Lieux / appellations
    'CHT':       'Château',
    'BDX':       'Bordeaux',
    'HT':        'Haut',
    'MEDOC':     'Médoc',
    'ST':        'Saint',
    'EMILION':   'Émilion',
    'CTES':      'Côtes',
    'CDR':       'Côtes du Rhône',
    # Cépage
    'CAB':       'Cabernet',
    'SAUV':      'Sauvignon',
    'CARB':      'Carménère',
    # Qualificatifs
    'GD':        'Grand',
    'PT':        'Petit',
    'PD':        'Pont-de',
    'BV':        'Baron',
    'RESERV':    'Réserve',
    'RSERV':     'Réserve',
    'RESERV.':   'Réserve',
    'SELEC':     'Sélection',
    'SPEC':      'Spéciale',
    'SPECIALE':  'Spéciale',
    'CUV':       'Cuvée',
    'CUVEE':     'Cuvée',
    'DOM':       'Domaine',
    'DOMAINE':   'Domaine',
    # Conditionnement
    'ASS':       'Assortiment',
    'BTLE':      'Bouteille',
    'PRIV':      'Privée',
    # Boissons
    'SING':      'Single',
    'ANS':       'ans',
    'BLCHE':     'Blanche',
    'APPERITIF': 'Apéritif',
    'APERITIF':  'Apéritif',
    'APERITIVO': 'Apéritivo',
    'BRUT':      'Brut',
    'DEMI-SEC':  'Demi-Sec',
    'WHISKY':    'Whisky',
    'COGNAC':    'Cognac',
    'COGNIAC':   'Cognac',
    'BRANDY':    'Brandy',
    'RHUM':      'Rhum',
    'VODKA':     'Vodka',
    'TEQUILA':   'Tequila',
    'LIQUEUR':   'Liqueur',
    'VERMOUTH':  'Vermouth',
    'VERMOUH':   'Vermouth',
    'VERMUTI':   'Vermouth',
    'GIN':       'Gin',
    'VIN':       'Vin',
    'VINS':      'Vins',
    'BLANC':     'Blanc',
    'ROUGE':     'Rouge',
    'ROSE':      'Rosé',
    'ROSÉ':      'Rosé',
    'MOUSSEUX':  'Mousseux',
    'PROSECCO':  'Prosecco',
    'CHAMPAGNE': 'Champagne',
    'PORTO':     'Porto',
    'KANNETTE':  'Cannette',
    'CANNETTE':  'Cannette',
    'N°':        'N°',
}

# Unités de volume : harmoniser la casse
VOLUME_RE = re.compile(
    r'(\d+(?:[.,]\d+)?)\s*(CL|ML|L)\b', re.IGNORECASE
)


def _normalize_volume(m):
    val = m.group(1).replace(',', '.')
    unit = m.group(2).lower()
    # Keep L uppercase for readability: 1L, 1.5L
    if unit == 'l':
        return f'{val}L'
    return f'{val}{unit}'


def expand_name(designation):
    """Retourne le nom complet sans abréviation, en titre."""
    text = str(designation)

    # Corriger les artefacts d'encodage courants
    text = text.replace('NÂ°', 'N°').replace('Â°', '°').replace('Â', '')

    # Normaliser les volumes
    text = VOLUME_RE.sub(_normalize_volume, text)

    # Appliquer les remplacements de phrases (multi-mots)
    for old, new in PHRASE_MAP:
        if callable(new):
            text = re.sub(old, new, text, flags=re.IGNORECASE)
        else:
            text = re.sub(r'\b' + re.escape(old) + r'\b', new,
                          text, flags=re.IGNORECASE)

    # Appliquer les remplacements mot à mot
    tokens = re.split(r'(\s+)', text)   # garder les espaces
    result = []
    # Ensemble des mots déjà présents (pour éviter les doublons)
    present_words = {w.lower() for w in text.split()}
    # Abréviations de type vin/boisson : on ne les ajoute pas si leur expansion
    # est déjà dans le texte
    type_abbrevs = {
        'MSX': 'mousseux', 'MOUSS': 'mousseux',
        'RGE': 'rouge', 'BLC': 'blanc', 'RSE': 'rosé',
        'MLX': 'moelleux', 'MOEL': 'moelleux', 'MOELL': 'moelleux',
    }
    for tok in tokens:
        if tok.strip() == '':
            result.append(tok)
            continue
        key = re.sub(r'[.,;:!?]$', '', tok.upper())
        if key in type_abbrevs and type_abbrevs[key] in present_words:
            # L'information est déjà exprimée : on saute cet abrégé
            continue
        if key in WORD_MAP:
            expanded = WORD_MAP[key]
            result.append(expanded)
            present_words.add(expanded.lower())
        else:
            # Title-case pour les tokens en MAJUSCULES (inclut lettres+ponct.)
            if tok.upper() == tok and any(c.isalpha() for c in tok):
                # Mettre en majuscule chaque séquence alphabétique
                result.append(re.sub(r'[A-Za-zÀ-öø-ÿ]+',
                                     lambda m2: m2.group(0).capitalize(), tok))
            else:
                result.append(tok)
    text = ' '.join(t for t in result if t.strip())

    return text.strip()


# ---------------------------------------------------------------------------
# Programme principal
# ---------------------------------------------------------------------------

def main():
    source_file = 'LISTE PRIX ARTICLE 2.xlsx'
    output_file = 'Articles_VIN_WHI_enriched.xlsx'

    print(f'Lecture de {source_file}…')
    df = pd.read_excel(source_file,
                       sheet_name="Liste d'articles (2)",
                       header=1)

    # Renommer proprement les colonnes
    df.columns = ['Référence', 'Désignation', 'Prix TTC']

    # Filtre VIN et WHI
    ref_upper = df['Référence'].astype(str).str.upper()
    mask = ref_upper.str.startswith('VIN') | ref_upper.str.startswith('WHI')
    df_filtered = df[mask].copy().reset_index(drop=True)
    print(f'{len(df_filtered)} articles sélectionnés (VIN + WHI).')

    # Ajout des nouvelles colonnes
    df_filtered['Contenance (cl)'] = df_filtered['Désignation'].apply(extract_contenance)
    df_filtered['Catégorie'] = df_filtered.apply(
        lambda r: get_categorie(r['Référence'], r['Désignation']), axis=1)
    df_filtered['Degré d\'alcool'] = df_filtered.apply(
        lambda r: get_degre_alcool(r['Catégorie'], r['Désignation']), axis=1)
    df_filtered['Nom'] = df_filtered['Désignation'].apply(expand_name)

    # Réorganiser les colonnes
    df_out = df_filtered[[
        'Référence',
        'Désignation',
        'Nom',
        'Contenance (cl)',
        'Catégorie',
        'Degré d\'alcool',
        'Prix TTC',
    ]]

    # Export Excel avec mise en forme
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_out.to_excel(writer, sheet_name='Articles VIN & WHI', index=False)

        wb  = writer.book
        ws  = writer.sheets['Articles VIN & WHI']

        # Formats
        header_fmt = wb.add_format({
            'bold': True, 'bg_color': '#1F4E79', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        })
        cell_fmt = wb.add_format({
            'border': 1, 'valign': 'vcenter'
        })
        price_fmt = wb.add_format({
            'border': 1, 'valign': 'vcenter', 'num_format': '#,##0'
        })
        center_fmt = wb.add_format({
            'border': 1, 'valign': 'vcenter', 'align': 'center'
        })

        # Largeurs des colonnes
        col_widths = [14, 55, 55, 16, 30, 16, 12]
        headers = list(df_out.columns)
        for i, (w, h) in enumerate(zip(col_widths, headers)):
            ws.set_column(i, i, w)
            ws.write(0, i, h, header_fmt)

        # Lignes de données
        for row_idx, row in df_out.iterrows():
            excel_row = row_idx + 1
            ws.write(excel_row, 0, row['Référence'],        cell_fmt)
            ws.write(excel_row, 1, row['Désignation'],      cell_fmt)
            ws.write(excel_row, 2, row['Nom'],              cell_fmt)
            # Contenance
            val = row['Contenance (cl)']
            ws.write(excel_row, 3, val if pd.notna(val) else '', center_fmt)
            ws.write(excel_row, 4, row['Catégorie'],        cell_fmt)
            ws.write(excel_row, 5, row['Degré d\'alcool'],  center_fmt)
            # Prix
            prix = row['Prix TTC']
            ws.write(excel_row, 6, prix if pd.notna(prix) else '', price_fmt)

        ws.set_row(0, 28)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df_out), len(df_out.columns) - 1)

    print(f'Fichier généré : {output_file}')

    # Statistiques
    print('\n=== Répartition par catégorie ===')
    print(df_filtered['Catégorie'].value_counts().to_string())

    print('\n=== Exemples de noms enrichis ===')
    sample = df_filtered[['Référence', 'Désignation', 'Nom',
                           'Contenance (cl)', 'Catégorie', 'Degré d\'alcool']].head(20)
    print(sample.to_string(index=False))


if __name__ == '__main__':
    main()
