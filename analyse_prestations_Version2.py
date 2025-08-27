import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

SOURCE_FILE = 'Classeur1.xlsx'  # Adapter si besoin
OUTPUT_EXCEL = 'rapport_prestations.xlsx'
OUTPUT_FIG = 'montant_par_acte.png'

def charger_donnees(fichier: str) -> pd.DataFrame:
    """Charge le fichier Excel et parse les dates si possible."""
    if not Path(fichier).exists():
        raise FileNotFoundError(f"Fichier introuvable: {fichier}")
    df_local = pd.read_excel(fichier, engine='openpyxl')
    # Harmoniser les noms de colonnes (strip + lower)
    df_local.columns = [c.strip() for c in df_local.columns]
    # Tenter conversion date
    if 'date' in [c.lower() for c in df_local.columns]:
        # Trouver le nom exact respectant la casse d'origine
        date_col = [c for c in df_local.columns if c.lower() == 'date'][0]
        df_local[date_col] = pd.to_datetime(df_local[date_col], errors='coerce')
    return df_local

def nettoyer(df: pd.DataFrame) -> pd.DataFrame:
    df = df.drop_duplicates().copy()
    # Normalisation nom centre
    centre_col_candidates = [c for c in df.columns if c.lower() in {'centre', 'centre_nom', 'centre name', 'centre_name'}]
    if centre_col_candidates:
        ccol = centre_col_candidates[0]
        df[ccol] = (df[ccol].astype(str).str.strip().str.upper()
                     .str.replace(' +', ' ', regex=True))
    # Filtrer années
    date_cols = [c for c in df.columns if c.lower() == 'date']
    if date_cols:
        dcol = date_cols[0]
        df = df[df[dcol].dt.year.isin([2024, 2025])]
    # Nettoyer montant (remplacer virgules milliers)
    montant_col_candidates = [c for c in df.columns if c.lower() in {'montant', 'montant demandé', 'montant_demande'}]
    if montant_col_candidates:
        mcol = montant_col_candidates[0]
        df[mcol] = (df[mcol].astype(str)
                              .str.replace('\u202f', '')  # espaces fines éventuelles
                              .str.replace(' ', '')
                              .str.replace(',', '')
                              .str.replace('\u00a0', '')
                              .str.replace('\.', '')  # si séparateur milliers
                              )
        # Si certains montants avaient des décimales séparées par un point, l'étape précédente les a retirées.
        # On tente une conversion safe.
        df[mcol] = pd.to_numeric(df[mcol], errors='coerce')
    return df

def calculer_indicateurs(df: pd.DataFrame) -> dict:
    montant_col_candidates = [c for c in df.columns if c.lower() in {'montant', 'montant demandé', 'montant_demande'}]
    if not montant_col_candidates:
        return {}
    mcol = montant_col_candidates[0]
    return {
        'montant_total': float(df[mcol].sum()),
        'montant_moyen': float(df[mcol].mean()),
        'montant_mediane': float(df[mcol].median()),
        'nb_lignes': int(len(df))
    }

def repartition_par_acte(df: pd.DataFrame) -> pd.Series:
    acte_col_candidates = [c for c in df.columns if c.lower() in {'acte', "type d'acte", 'type'}]
    montant_col_candidates = [c for c in df.columns if c.lower() in {'montant', 'montant demandé', 'montant_demande'}]
    if not acte_col_candidates or not montant_col_candidates:
        return pd.Series(dtype=float)
    acol = acte_col_candidates[0]
    mcol = montant_col_candidates[0]
    return df.groupby(acol)[mcol].sum().sort_values(ascending=False)

def tracer(par_acte: pd.Series):
    if par_acte.empty:
        print("Aucune donnée pour tracer le graphique.")
        return
    plt.figure(figsize=(10, 5))
    par_acte.plot(kind='bar', color='#2E86C1')
    plt.title('Montant total par acte')
    plt.ylabel('Montant (somme)')
    plt.xlabel('Acte')
    plt.tight_layout()
    plt.savefig(OUTPUT_FIG)
    plt.close()
    print(f"Graphique sauvegardé: {OUTPUT_FIG}")

def exporter(df: pd.DataFrame, par_acte: pd.Series, indicateurs: dict):
    with pd.ExcelWriter(OUTPUT_EXCEL) as writer:
        df.to_excel(writer, sheet_name='Donnees_nettoyees', index=False)
        if not par_acte.empty:
            par_acte.to_frame(name='montant_total').to_excel(writer, sheet_name='Repartition_actes')
        if indicateurs:
            (pd.Series(indicateurs)
               .rename('valeur')
               .to_frame()
               .to_excel(writer, sheet_name='Indicateurs'))
    print(f"Fichier Excel généré: {OUTPUT_EXCEL}")

def main():
    try:
        df = charger_donnees(SOURCE_FILE)
        df = nettoyer(df)
        indicateurs = calculer_indicateurs(df)
        if indicateurs:
            print('Indicateurs:')
            for k, v in indicateurs.items():
                print(f"  - {k}: {v:,.0f}" if isinstance(v, float) else f"  - {k}: {v}")
        serie_acte = repartition_par_acte(df)
        if not serie_acte.empty:
            print('\nTop 5 actes:')
            print(serie_acte.head(5))
        tracer(serie_acte)
        exporter(df, serie_acte, indicateurs)
    except Exception as e:
        print(f"Erreur lors du traitement: {e}")

if __name__ == '__main__':
    main()