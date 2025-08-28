"""Analyse complète des prestations Mutuelle de Santé.
Assumptions:
- Fichier source: Classeur1.xlsx (feuille active par défaut)
- Colonnes attendues (selon échantillon):
  adherent_code, adherent_nom, adherent_prenom, adherent_genre,
  beneficiaire, identifiant_prestation, acte, date, type, sous_type,
  montant, validite, centre_nom
- Colonnes manquantes du plan (region/province, montant payé) sont absentes.
  -> On suppose region/province indisponible pour l'instant.
  -> On suppose montant_paye == montant quand validite == 'accepté'.
  -> Statut = validite.
"""
from __future__ import annotations
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import base64
from io import BytesIO
import math
from datetime import datetime
from matplotlib.patches import Rectangle
import time

try:
    from fpdf import FPDF
    _HAS_FPDF = True
except ImportError:  # fpdf2 non installé
    _HAS_FPDF = False

SOURCE_FILE = 'Classeur1.xlsx'
OUTPUT_EXCEL = 'rapport_prestations_complet.xlsx'
OUTPUT_PDF = 'rapport_prestations.pdf'
FIG_DIR = Path('figures')
FIG_DIR.mkdir(exist_ok=True)
LOGO_PATH = Path('logo_mupol.png')  # placer votre logo ici (PNG/JPG)
# Sécurité PDF (laisser vide pour désactiver)
PDF_USER_PWD = 'Lecture2025'       # mot de passe utilisateur (ouverture)
PDF_OWNER_PWD = 'AdminSecure#2025' # mot de passe propriétaire (permissions)
PDF_PERMISSIONS = ['print']        # ex: ['print', 'copy']

# Stockage global des informations sur les doublons pour reporting
DUPLICATES_INFO: dict = {}
# Lignes dont les dates ont été corrigées (postérieures à juillet 2025 -> année 2024)
DATES_CORRIGEES_DF: pd.DataFrame | None = None

# Style global
sns.set_theme(style='whitegrid', context='talk', palette='Set2')
plt.rcParams.update({
    'figure.dpi': 110,
    'axes.titlesize': 14,
    'axes.labelsize': 12,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'grid.color': '#d1d1d1',  # légèrement plus visible
    'grid.linewidth': 0.3,     # plus fin
    'grid.linestyle': '-',
    'axes.grid': True,
    'axes.grid.axis': 'y'
})

# ------------------ Chargement & Préparation ------------------ #

def charger(fichier: str) -> pd.DataFrame:
    if not Path(fichier).exists():
        raise FileNotFoundError(f"Fichier introuvable: {fichier}")
    df = pd.read_excel(fichier, engine='openpyxl')
    # Normaliser noms colonnes (strip)
    df.columns = [c.strip() for c in df.columns]
    # Date -> datetime
    date_cols = [c for c in df.columns if c.lower() == 'date']
    if date_cols:
        dc = date_cols[0]
        df[dc] = pd.to_datetime(df[dc], errors='coerce')
        # Correction demandée : toute date postérieure à juillet 2025 -> année forcée à 2024 (mois & jour conservés)
        try:
            global DATES_CORRIGEES_DF
            mask_fix = (df[dc].dt.year == 2025) & (df[dc].dt.month > 7)  # août (8) à décembre (12) 2025
            if mask_fix.any():
                subset = df.loc[mask_fix].copy()
                subset['date_originale'] = subset[dc]
                subset['date_corrigee'] = subset[dc].apply(lambda d: d.replace(year=2024) if pd.notna(d) else d)
                # Appliquer correction sur le dataframe principal
                df.loc[mask_fix, dc] = subset['date_corrigee']
                # Sauvegarde globale (toutes colonnes + 2 colonnes info)
                cols_order = ['date_originale','date_corrigee'] + [c for c in subset.columns if c not in {'date_originale','date_corrigee'}]
                DATES_CORRIGEES_DF = subset[cols_order]
        except Exception:
            pass
    # Montant -> numérique (suppression séparateurs)
    montant_cols = [c for c in df.columns if c.lower() in {'montant', 'montant demande', 'montant demandé'}]
    if montant_cols:
        mc = montant_cols[0]
        df[mc] = (df[mc].astype(str)
                        .str.replace('\u202f', '')
                        .str.replace('\u00a0', '')
                        .str.replace(' ', '')
                        .str.replace(',', '')
                        )
        df[mc] = pd.to_numeric(df[mc], errors='coerce')
    # Centre
    centre_cols = [c for c in df.columns if 'centre' in c.lower()]
    if centre_cols:
        cc = centre_cols[0]
        df[cc] = (df[cc].astype(str).str.strip().str.upper()
                              .str.replace(' +', ' ', regex=True))
    # Statut -> statut
    if 'validite' in [c.lower() for c in df.columns]:
        col = [c for c in df.columns if c.lower() == 'validite'][0]
        df.rename(columns={col: 'statut'}, inplace=True)
        df['statut'] = df['statut'].str.strip().str.lower()
    # Vérification et capture des doublons avant suppression
    global DUPLICATES_INFO
    # Doublons ligne entière
    dups_full = df[df.duplicated(keep=False)].copy()
    # Détection d'un identifiant unique potentiel
    id_col = next((c for c in df.columns if 'identifiant' in c.lower()), None)
    if id_col:
        dups_id = df[df.duplicated(subset=[id_col], keep=False)].copy()
    else:
        dups_id = pd.DataFrame()
    # Clé composite fréquente (si colonnes présentes)
    composite_cols = [c for c in ['adherent_code','date','acte','montant'] if c in df.columns]
    if composite_cols:
        dups_composite = df[df.duplicated(subset=composite_cols, keep=False)].copy()
    else:
        dups_composite = pd.DataFrame()
    DUPLICATES_INFO = {
        'doublons_lignes': dups_full,
        'doublons_identifiant': dups_id,
        'doublons_cle_composite': dups_composite,
        'nb_doublons_lignes': int(dups_full.shape[0]),
        'nb_doublons_identifiant': int(dups_id.shape[0]) if not dups_id.empty else 0,
        'nb_doublons_cle_composite': int(dups_composite.shape[0]) if not dups_composite.empty else 0
    }
    # Suppression doublons lignes pour analyses
    df.drop_duplicates(inplace=True)
    # Filtre années 2024-2025 si date existe
    if date_cols:
        dc = date_cols[0]
        df = df[df[dc].dt.year.isin([2024, 2025])]
    return df

# ------------------ Indicateurs ------------------ #

def compute_global(df: pd.DataFrame) -> dict:
    montant_col = next((c for c in df.columns if c.lower().startswith('montant')), None)
    date_col = next((c for c in df.columns if c.lower() == 'date'), None)
    adherent_col = next((c for c in df.columns if 'adherent_code' in c.lower()), None)
    d = {}
    if montant_col:
        d['montant_total'] = float(df[montant_col].sum())
        d['montant_moyen'] = float(df[montant_col].mean()) if len(df) else math.nan
        d['montant_mediane'] = float(df[montant_col].median()) if len(df) else math.nan
    d['nb_prestations'] = int(len(df))
    if adherent_col:
        d['nb_mutualistes_distincts'] = int(df[adherent_col].nunique())
    if 'statut' in df.columns:
        total = len(df)
        acceptes = (df['statut'] == 'accepté').sum() + (df['statut'] == 'accepte').sum()
        d['taux_acceptation_pct'] = (acceptes / total * 100) if total else math.nan
        if montant_col:
            montant_total = float(df[montant_col].sum())
            montant_paye = float(df.loc[df['statut'].isin(['accepté','accepte']), montant_col].sum())
            d['montant_paye_total'] = montant_paye
            d['montant_non_paye'] = montant_total - montant_paye
            d['pourcentage_paye_pct'] = (montant_paye / montant_total * 100) if montant_total else math.nan
    if montant_col and adherent_col and 'taux_acceptation_pct' in d:
        d['cout_moyen_mutualiste'] = float(df[montant_col].sum() / df[adherent_col].nunique()) if df[adherent_col].nunique() else math.nan
    # Période couverte
    if date_col:
        d['periode_min'] = df[date_col].min()
        d['periode_max'] = df[date_col].max()
    return d

# ------------------ Libellés Français ------------------ #

INDIC_LABELS = {
    'montant_total': 'Montant total des prestations',
    'montant_moyen': 'Montant moyen par prestation',
    'montant_mediane': 'Montant médian par prestation',
    'nb_prestations': 'Nombre de prestations',
    'nb_mutualistes_distincts': 'Nombre de mutualistes distincts',
    'taux_acceptation_pct': "Taux d'acceptation (%)",
    'cout_moyen_mutualiste': 'Coût moyen par mutualiste',
    'montant_paye_total': 'Montant total payé (accepté)',
    'montant_non_paye': 'Montant non payé',
    'pourcentage_paye_pct': 'Pourcentage payé (%)',
    'periode_min': 'Date première prestation',
    'periode_max': 'Date dernière prestation'
}

COL_LABELS = {
    'sum': 'Montant total',
    'count': 'Nombre de prestations',
    'mean': 'Montant moyen par prestation',
    'mutualistes_distincts': 'Nombre de mutualistes distincts'
}

HEADER_FIRST_MAP = {
    'repartition_par_acte': 'Acte',
    'repartition_par_centre': 'Centre de santé',
    'repartition_par_region': 'Région',
    'repartition_par_province': 'Province',
    'repartition_par_type': 'Type',
    'repartition_par_sous_type': 'Sous-type',
    'repartition_par_statut': 'Statut',
    'repartition_par_partenaire': 'Partenaire'
}

# ------------------ Répartitions ------------------ #

def repartitions(df: pd.DataFrame) -> dict[str, pd.DataFrame | pd.Series]:
    out = {}
    montant_col = next((c for c in df.columns if c.lower().startswith('montant')), None)
    adherent_col = next((c for c in df.columns if 'adherent_code' in c.lower()), None)
    if not montant_col:
        return out
    def agg(col):
        base = df.groupby(col)[montant_col].agg(['sum', 'count', 'mean'])
        if adherent_col and adherent_col in df.columns:
            mut = df.groupby(col)[adherent_col].nunique().rename('mutualistes_distincts')
            base = base.join(mut)
        return base.sort_values('sum', ascending=False)
    mapping = {
        'repartition_par_acte': 'acte',
        'repartition_par_type': 'type',
        'repartition_par_sous_type': 'sous_type',
        'repartition_par_centre': next((c for c in df.columns if 'centre' in c.lower()), None),
        'repartition_par_partenaire': next((c for c in df.columns if 'parten' in c.lower()), None),
        'repartition_par_region': next((c for c in df.columns if c.lower() == 'region'), None),
        'repartition_par_province': None,  # géré séparément via multi-index région+province
        'repartition_par_statut': 'statut'
    }
    for key, col in mapping.items():
        if key not in ('repartition_par_province',) and col and col in df.columns:
            out[key] = agg(col)
    # Province multi-index region+province
    region_col = mapping['repartition_par_region']
    province_col = next((c for c in df.columns if c.lower() == 'province'), None)
    if region_col and province_col and region_col in df.columns and province_col in df.columns:
        grp = df.groupby([region_col, province_col])
        base = grp[montant_col].agg(['sum', 'count', 'mean'])
        if adherent_col and adherent_col in df.columns:
            mut = grp[adherent_col].nunique().rename('mutualistes_distincts')
            base = base.join(mut)
        reg_order = base.groupby(level=0)['sum'].sum().sort_values(ascending=False).index
        ordered = []
        for reg in reg_order:
            sub = base.loc[reg].sort_values('sum', ascending=False)
            for prov, row in sub.iterrows():
                ordered.append((reg, prov, *row.values))
        cols = ['Region','Province'] + list(base.columns)
        out['repartition_par_province'] = pd.DataFrame(ordered, columns=cols)
    # Top centres
    if 'repartition_par_centre' in out:
        out['top10_centres'] = out['repartition_par_centre'].head(10)
    return out

# ------------------ Analyse temporelle ------------------ #

def analyse_temporelle(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    date_col = next((c for c in df.columns if c.lower() == 'date'), None)
    montant_col = next((c for c in df.columns if c.lower().startswith('montant')), None)
    if not date_col or not montant_col:
        return pd.DataFrame(), pd.DataFrame()
    temp = df[[date_col, montant_col]].copy()
    temp['annee'] = temp[date_col].dt.year
    temp['mois'] = temp[date_col].dt.to_period('M').astype(str)
    evo = temp.groupby('mois')[montant_col].agg(['sum', 'count']).rename(columns={'sum':'montant_total','count':'nb_prestations'})
    # Comparaison T1
    temp['trimestre'] = temp[date_col].dt.to_period('Q').astype(str)
    t1 = temp[temp[date_col].dt.quarter == 1]
    comp_t1 = t1.groupby('annee')[montant_col].agg(['sum', 'count']).rename(columns={'sum':'montant_total','count':'nb_prestations'})
    return evo, comp_t1

# ------------------ Graphiques ------------------ #

def _annotate_bars(ax):
    for p in ax.patches:
        height = p.get_height()
        if not math.isnan(height):
            ax.annotate(f"{int(height):,}".replace(',', ' '),
                        (p.get_x() + p.get_width()/2, height),
                        ha='center', va='bottom', fontsize=9, rotation=0)

def plot_bar(df_series: pd.Series, title: str, filename: str, top: int | None = None, horizontal=False, pie=False, counts_series: pd.Series | None = None):
    if df_series.empty:
        return None
    data = df_series.head(top) if top else df_series
    fig, ax = plt.subplots(figsize=(10,5))
    if pie:
        total = data.sum()
        percentages = (data / total * 100).round(1)
        wedges, texts = ax.pie(data.values, startangle=90)
        # Pas d'autopct pour éviter chevauchement; légende externe
        legend_labels = [f"{idx} – {val:,.0f} ({pct}%)".replace(',', ' ') for idx, val, pct in zip(data.index, data.values, percentages)]
        ax.legend(wedges, legend_labels, title='Actes', loc='center left', bbox_to_anchor=(1.0, 0.5), fontsize=8)
        ax.set_ylabel('')
    else:
        if horizontal:
            sns.barplot(x=data.values, y=data.index, ax=ax)
            ax.set_xlabel('Montant total')
            ax.set_ylabel('')
            for i, (idx, v) in enumerate(zip(data.index, data.values)):
                if counts_series is not None and idx in counts_series.index:
                    c = counts_series.loc[idx]
                    ax.text(v, i, f" {int(v):,} ({int(c)})".replace(',', ' '), va='center', fontsize=9)
                else:
                    ax.text(v, i, f" {int(v):,}".replace(',', ' '), va='center', fontsize=9)
        else:
            sns.barplot(x=data.index, y=data.values, ax=ax)
            ax.set_ylabel('Montant total')
            ax.set_xlabel('')
            # Annotations sur chaque barre
            for p, idx in zip(ax.patches, data.index):
                height = p.get_height()
                if not math.isnan(height):
                    if counts_series is not None and idx in counts_series.index:
                        c = counts_series.loc[idx]
                        txt = f"{int(height):,} ({int(c)})".replace(',', ' ')
                    else:
                        txt = f"{int(height):,}".replace(',', ' ')
                    ax.annotate(txt, (p.get_x() + p.get_width()/2, height), ha='center', va='bottom', fontsize=9)
        ax.tick_params(axis='x', rotation=35 if not horizontal else 0)
    ax.set_title(title)
    plt.tight_layout()
    # Bordure carrée
    rect = Rectangle((0.005,0.005),0.99,0.99, transform=fig.transFigure, fill=False, lw=1.4, edgecolor='#333')
    fig.patches.append(rect)
    path = FIG_DIR / filename
    fig.savefig(path)
    plt.close(fig)
    return path

def plot_pie_group_small(series: pd.Series, title: str, filename: str, threshold: float = 0.03, donut: bool = True):
    if series.empty:
        return None
    s = series.dropna().copy()
    total = s.sum()
    if total == 0:
        return None
    props = s / total
    small = props < threshold
    if small.sum() > 1:
        s_other = s[small].sum()
        s = s[~small]
        s.loc['Autres'] = s_other
    s = s.sort_values(ascending=False)
    fig, ax = plt.subplots(figsize=(9.5,6))
    pct = (s / s.sum() * 100).round(1)
    wedgeprops = {'width':0.5} if donut else None
    wedges, _ = ax.pie(s.values, startangle=90, wedgeprops=wedgeprops, labels=None)
    legend_labels = [f"{idx} – {val:,.0f} ({p:.1f}%)".replace(',', ' ') for idx, val, p in zip(s.index, s.values, pct)]
    ax.legend(wedges, legend_labels, title='Catégories', loc='center left', bbox_to_anchor=(1.0, 0.5), fontsize=8)
    ax.set_ylabel('')
    ax.set_title(title)
    plt.tight_layout()
    # Bordure carrée
    rect = Rectangle((0.005,0.005),0.99,0.99, transform=fig.transFigure, fill=False, lw=1.4, edgecolor='#333')
    fig.patches.append(rect)
    path = FIG_DIR / filename
    fig.savefig(path)
    plt.close(fig)
    return path

def plot_region_distribution(reps: dict) -> Path | None:
    """Barres horizontales montant par région avec nombre entre parenthèses (sans courbe)."""
    key = 'repartition_par_region'
    if key not in reps:
        return None
    df = reps[key]
    if not {'sum','count'}.issubset(df.columns):
        return None
    data = df.sort_values('sum', ascending=False)
    fig, ax = plt.subplots(figsize=(10, 0.45*len(data)+1.2))
    y = range(len(data))
    ax.barh(y, data['sum'], color='#76b5c5', alpha=0.85)
    ax.set_yticks(list(y))
    ax.set_yticklabels(data.index)
    ax.invert_yaxis()  # plus grand en haut
    ax.set_xlabel('Montant total')
    ax.set_ylabel('Région')
    for yi, (v, n) in enumerate(zip(data['sum'], data['count'])):
        val_str = f"{int(v):,}".replace(',', ' ')
        ax.text(v, yi, f" {val_str} ({int(n)})", va='center', fontsize=8)
    ax.set_title('Répartition par région – Montant (Nombre)')
    fig.tight_layout()
    rect = Rectangle((0.005,0.005),0.99,0.99, transform=fig.transFigure, fill=False, lw=1.4, edgecolor='#333')
    fig.patches.append(rect)
    path = FIG_DIR / 'repartition_region_montant_nombre.png'
    fig.savefig(path)
    plt.close(fig)
    return path

def plot_province_distribution(reps: dict, top: int = 20) -> Path | None:
    """Barres horizontales provinces (top N) montant + annotations nombre."""
    key = 'repartition_par_province'
    if key not in reps:
        return None
    df = reps[key].copy()
    if 'Province' not in df.columns or 'sum' not in df.columns or 'count' not in df.columns:
        return None
    df_top = df.sort_values('sum', ascending=False).head(top)
    fig, ax = plt.subplots(figsize=(11, 0.45*len(df_top)+1.5))
    y = range(len(df_top))
    ax.barh(y, df_top['sum'], color='#a5d296', alpha=0.9)
    ax.set_yticks(list(y))
    ax.set_yticklabels(df_top['Province'])
    ax.invert_yaxis()  # plus grand en haut
    ax.set_xlabel('Montant total')
    ax.set_title(f'Repartition par province – Top {len(df_top)} (Montant (Nombre))')
    for yi, (v, n) in enumerate(zip(df_top['sum'], df_top['count'])):
        val_str = f"{int(v):,}".replace(',', ' ')
        ax.text(v, yi, f" {val_str} ({int(n)})", va='center', fontsize=8)
    fig.tight_layout()
    rect = Rectangle((0.005,0.005),0.99,0.99, transform=fig.transFigure, fill=False, lw=1.4, edgecolor='#333')
    fig.patches.append(rect)
    path = FIG_DIR / 'repartition_province_montant_nombre.png'
    fig.savefig(path)
    plt.close(fig)
    return path

def annotate_line_no_overlap(ax, x_vals, y_vals, texts, min_gap_frac=0.025, x_offset=0.35):
    """Annotate line points with labels placed to the right, clustering to avoid vertical overlap.
    - min_gap_frac: minimal vertical gap between labels as fraction of data y-range.
    - x_offset: horizontal offset (data units) to the right of each point.
    The function expands y-limits if needed and draws connector lines for shifted labels."""
    if not x_vals:
        return
    y_min, y_max = float(min(y_vals)), float(max(y_vals))
    y_range = (y_max - y_min) or 1.0
    min_gap = y_range * min_gap_frac
    # Build clusters of indices whose y are within min_gap when sorted
    sorted_idx = sorted(range(len(y_vals)), key=lambda i: y_vals[i])
    clusters = []
    current = [sorted_idx[0]]
    for idx in sorted_idx[1:]:
        if abs(y_vals[idx] - y_vals[current[-1]]) <= min_gap:
            current.append(idx)
        else:
            clusters.append(current)
            current = [idx]
    clusters.append(current)
    adjusted_y = [None]*len(y_vals)
    for cluster in clusters:
        if len(cluster) == 1:
            adjusted_y[cluster[0]] = y_vals[cluster[0]]
            continue
        # Spread cluster labels vertically centered around their mean
        orig_vals = [y_vals[i] for i in cluster]
        center = sum(orig_vals)/len(orig_vals)
        span = min_gap * (len(cluster)-1)
        start = center - span/2
        for pos, i in enumerate(cluster):
            adjusted_y[i] = start + pos*min_gap
    # Possibly extend y-limits if we pushed above
    new_y_max = max(adjusted_y + [y_max])
    if new_y_max > y_max:
        ax.set_ylim(top=new_y_max + min_gap)
    # Adjust x-limits to give space on right
    x_max = max(x_vals)
    ax.set_xlim(-0.5, x_max + 1.2)  # extra room
    # Draw annotations
    for i, (x, orig_y, lab_y, txt) in enumerate(zip(x_vals, y_vals, adjusted_y, texts)):
        ax.annotate(txt, (x + x_offset, lab_y), ha='left', va='center', fontsize=8)
        if abs(lab_y - orig_y) > 1e-9 or x_offset != 0:
            ax.plot([x, x + x_offset*0.9], [orig_y, lab_y], color='gray', linewidth=0.5, linestyle=':')

def format_abbrev(val: float) -> str:
    """Format large numbers with French style abbreviations (k, M, Md)."""
    if pd.isna(val):
        return ''
    abs_v = abs(val)
    if abs_v >= 1_000_000_000:
        return f"{val/1_000_000_000:.2f} Md".replace('.', ',')
    if abs_v >= 1_000_000:
        return f"{val/1_000_000:.2f} M".replace('.', ',')
    if abs_v >= 10_000:
        return f"{val/1_000:.1f} k".replace('.', ',')
    return f"{int(val):,}".replace(',', ' ')
def annotate_line_no_overlap(ax, x_vals, y_vals, texts, min_gap_frac=0.02, base_x_offset=0.35):
    """Advanced annotation: multi-line, adaptive staggering horizontally & vertically.
    texts: list of strings (can contain '\n')."""
    if not x_vals:
        return
    y_min, y_max = float(min(y_vals)), float(max(y_vals))
    y_range = (y_max - y_min) or 1.0
    min_gap = y_range * min_gap_frac
    # Sort points by y
    order = sorted(range(len(y_vals)), key=lambda i: y_vals[i])
    adjusted_y = [None]*len(y_vals)
    clusters = []
    cluster = [order[0]]
    for i in order[1:]:
        if abs(y_vals[i] - y_vals[cluster[-1]]) <= min_gap:
            cluster.append(i)
        else:
            clusters.append(cluster)
            cluster = [i]
    clusters.append(cluster)
    for cl in clusters:
        if len(cl) == 1:
            adjusted_y[cl[0]] = y_vals[cl[0]]
        else:
            center = sum(y_vals[k] for k in cl)/len(cl)
            span = min_gap * max(1, (len(cl)-1)) * 1.25
            start = center - span/2
            step = span / (len(cl)-1)
            for idx,pos in enumerate(cl):
                adjusted_y[pos] = start + idx*step
    # Extend ylim if necessary
    new_top = max(adjusted_y + [y_max])
    if new_top > y_max:
        ax.set_ylim(top=new_top + min_gap*0.8)
    x_max = max(x_vals)
    ax.set_xlim(-0.5, x_max + 2.2)
    # Draw annotations
    for cl in clusters:
        for idx, point_idx in enumerate(cl):
            x = x_vals[point_idx]
            oy = y_vals[point_idx]
            ly = adjusted_y[point_idx]
            # Horizontal offset with slight incremental shift to create a fan
            local_off = base_x_offset + (idx * 0.15 if len(cl) > 2 else idx * 0.1)
            ax.annotate(texts[point_idx], (x + local_off, ly), ha='left', va='center', fontsize=8, linespacing=1.0)
            if abs(ly - oy) > 1e-9 or local_off != 0:
                ax.plot([x, x + local_off*0.85], [oy, ly], color='gray', linewidth=0.5, linestyle=':')

# ------------------ Export Excel ------------------ #

def export_excel(global_indic: dict, reps: dict, evo: pd.DataFrame, comp_t1: pd.DataFrame, df: pd.DataFrame):
    # Préparation indicateurs avec labels FR
    indic_fr = {INDIC_LABELS.get(k, k): v for k, v in global_indic.items()}
    target = OUTPUT_EXCEL
    attempt = 0
    while True:
        try:
            writer_ctx = pd.ExcelWriter(target)
            break
        except PermissionError:
            attempt += 1
            if attempt > 1:
                # Générer un nouveau nom horodaté
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                target = f"rapport_prestations_complet_{ts}.xlsx"
            time.sleep(0.4)
    with writer_ctx as writer:
        (pd.Series(indic_fr).to_frame('Valeur')).to_excel(writer, sheet_name='Indicateurs')
        for key, val in reps.items():
            sheet = key.replace('repartition_', 'rep_')[:31]
            df_tmp = val.copy()
            if key != 'repartition_par_province' and not isinstance(df_tmp.index, pd.MultiIndex):
                df_tmp = df_tmp.reset_index()
                # Ajout de la colonne Numéro sauf pour le tableau statut
                if key != 'repartition_par_statut':
                    df_tmp.insert(0, 'Numéro', range(1, len(df_tmp) + 1))
                if len(df_tmp.columns):
                    first_col = df_tmp.columns[1] if 'Numéro' in df_tmp.columns else df_tmp.columns[0]
                    # Renommer première colonne de catégorie
                    df_tmp.rename(columns={first_col: HEADER_FIRST_MAP.get(key, first_col)}, inplace=True)
            elif key == 'repartition_par_province':
                # Ajouter Numéro en première colonne (sauf statut exclu déjà géré)
                if len(df_tmp):
                    df_tmp.insert(0, 'Numéro', range(1, len(df_tmp) + 1))
                    # Réordonner pour avoir Numéro, Province, Région puis métriques (Province avant Région comme dans HTML)
                    cols = df_tmp.columns.tolist()
                    order = ['Numéro']
                    if 'Province' in cols:
                        order.append('Province')
                    if 'Region' in cols:
                        order.append('Region')
                    for c in cols:
                        if c not in order:
                            order.append(c)
                    df_tmp = df_tmp[order]
            df_tmp.rename(columns=COL_LABELS, inplace=True)
            df_tmp.to_excel(writer, sheet_name=sheet, index=False)
        if not evo.empty:
            evo_fr = evo.rename(columns={'montant_total': 'Montant total', 'nb_prestations': 'Nombre de prestations'})
            evo_fr.to_excel(writer, sheet_name='Evolution_mensuelle')
        if not comp_t1.empty:
            comp_fr = comp_t1.rename(columns={'montant_total': 'Montant total', 'nb_prestations': 'Nombre de prestations'})
            comp_fr.to_excel(writer, sheet_name='Comparaison_T1')
        # Données brutes
        df.to_excel(writer, sheet_name='Donnees_nettoyees', index=False)
        # Doublons (reporting)
        if DUPLICATES_INFO:
            for key_sheet, df_dup in [
                ('Doublons_lignes', DUPLICATES_INFO.get('doublons_lignes')),
                ('Doublons_identifiant', DUPLICATES_INFO.get('doublons_identifiant')),
                ('Doublons_cle_comp', DUPLICATES_INFO.get('doublons_cle_composite'))
            ]:
                if isinstance(df_dup, pd.DataFrame) and not df_dup.empty:
                    # Ajouter une colonne Occurrences si pertinent (pour identifiant / clé composite)
                    if key_sheet != 'Doublons_lignes':
                        col_sub = []
                        if key_sheet == 'Doublons_identifiant':
                            col_sub = [c for c in df_dup.columns if 'identifiant' in c.lower()][:1]
                        elif key_sheet == 'Doublons_cle_comp':
                            col_sub = [c for c in ['adherent_code','date','acte','montant'] if c in df_dup.columns]
                        if col_sub:
                            occ = (df_dup.groupby(col_sub)
                                         .size()
                                         .reset_index(name='Occurrences')
                                         .sort_values('Occurrences', ascending=False))
                            occ.to_excel(writer, sheet_name=(key_sheet + '_occ')[:31], index=False)
                    df_dup.to_excel(writer, sheet_name=key_sheet[:31], index=False)
        # Lignes de dates corrigées
        if DATES_CORRIGEES_DF is not None and not DATES_CORRIGEES_DF.empty:
            tmp_dates = DATES_CORRIGEES_DF.copy()
            # Formatage lisible dates
            for dc_col in ['date_originale','date_corrigee']:
                if dc_col in tmp_dates.columns:
                    tmp_dates[dc_col] = tmp_dates[dc_col].dt.strftime('%d/%m/%Y')
            tmp_dates.to_excel(writer, sheet_name='Dates_corrigees', index=False)
    if target != OUTPUT_EXCEL:
        print(f"[INFO] Fichier Excel initial verrouillé. Export sauvegardé sous: {target}")

# ------------------ Export PDF ------------------ #

def export_pdf(global_indic: dict, reps: dict, images_paths, evo: pd.DataFrame, comp_t1: pd.DataFrame):
    if not _HAS_FPDF:
        print("Bibliothèque fpdf2 non installée -> export PDF ignoré.")
        return
    pdf = FPDF()
    # Appliquer protection si mots de passe définis
    try:
        if PDF_USER_PWD and PDF_OWNER_PWD and hasattr(pdf, 'set_protection'):
            pdf.set_protection(PDF_PERMISSIONS, user_pwd=PDF_USER_PWD, owner_pwd=PDF_OWNER_PWD)
        elif PDF_USER_PWD and PDF_OWNER_PWD and not hasattr(pdf, 'set_protection'):
            print("Info: méthode set_protection indisponible dans cette version de fpdf2 -> PDF non chiffré.")
    except Exception as e:
        print(f"Avertissement: échec application protection PDF ({e})")
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    # Charger une police Unicode (DejaVuSans) si disponible, sinon fallback Helvetica
    unicode_font_path = None
    for candidate in [
        Path('DejaVuSans.ttf'),
        Path('C:/Windows/Fonts/DejaVuSans.ttf'),
        Path('C:/Windows/Fonts/DejaVuSansCondensed.ttf'),
        Path('C:/Windows/Fonts/arial.ttf')
    ]:
        if candidate.exists():
            unicode_font_path = candidate
            break
    if unicode_font_path:
        try:
            pdf.add_font('DejaVu', '', str(unicode_font_path), uni=True)
            pdf.add_font('DejaVu', 'B', str(unicode_font_path), uni=True)
            base_font = 'DejaVu'
        except Exception:
            base_font = 'Helvetica'
    else:
        base_font = 'Helvetica'
    pdf.set_font(base_font, 'B', 16)
    _sanitize = lambda s: str(s).replace('’', "'")
    pdf.cell(0, 10, _sanitize('Rapport Analytique des Prestations'), ln=1)
    # Logo (si présent) en haut à droite
    if LOGO_PATH.exists():
        try:
            pdf.image(str(LOGO_PATH), x=170, y=10, w=30)
        except Exception:
            pass
    pdf.set_font(base_font, '', 10)
    pdf.cell(0, 6, _sanitize(f"Généré le {datetime.now().strftime('%d/%m/%Y %H:%M')}"), ln=1)
    # Objectifs de l'analyse
    pdf.set_font(base_font, 'B', 12)
    pdf.cell(0, 8, _sanitize("Objectif de l'analyse"), ln=1)
    pdf.set_font(base_font, '', 9)
    objectifs = [
        _sanitize("Évaluer les dépenses de santé prises en charge par la mutuelle."),
        _sanitize("Comparer la répartition par actes, centres, partenaires et zones géographiques."),
        _sanitize("Mesurer la performance du remboursement (taux d'acceptation, paiements effectués)."),
        _sanitize("Identifier les tendances pour anticiper la charge financière future.")
    ]
    # Rendu manuel des puces (évite erreur largeur sur certains environnements)
    max_w = pdf.w - pdf.l_margin - pdf.r_margin
    for o in objectifs:
        words = o.split()
        line = "- "
        for w in words:
            test = line + w + ' '
            if pdf.get_string_width(test) > max_w:
                pdf.cell(0, 5, line.rstrip(), ln=1)
                line = "  " + w + ' '
            else:
                line = test
        if line.strip():
            pdf.cell(0, 5, line.rstrip(), ln=1)
    pdf.ln(2)
    # Indicateurs globaux
    pdf.set_font(base_font, 'B', 12)
    pdf.cell(0, 8, '1. Indicateurs globaux', ln=1)
    pdf.set_font(base_font, '', 9)
    for k, v in global_indic.items():
        label = INDIC_LABELS.get(k, k)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            if 'taux' in k:
                val_str = f"{v:.1f}%"
            else:
                val_str = f"{v:,.0f}".replace(',', ' ')
        elif isinstance(v, pd.Timestamp):
            val_str = v.strftime('%d/%m/%Y') if not pd.isna(v) else ''
        else:
            val_str = str(v)
        pdf.cell(95, 6, label, border=1)
        pdf.cell(95, 6, val_str, border=1, ln=1)
    # Répartitions clefs
    pdf.add_page()
    pdf.set_font('Helvetica', 'B', 12)
    pdf.cell(0, 8, '2. Répartitions', ln=1)
    pdf.set_font('Helvetica', '', 9)
    header_first_map = {
        'repartition_par_acte': 'Acte',
        'repartition_par_centre': 'Centre de santé',
        'repartition_par_region': 'Région',
        'repartition_par_province': 'Province',  # géré séparément (Province + Région)
        'repartition_par_type': 'Type',
        'repartition_par_sous_type': 'Sous-type',
        'repartition_par_statut': 'Statut'
    }
    for key in ['repartition_par_acte', 'repartition_par_centre', 'repartition_par_region', 'repartition_par_province', 'repartition_par_type', 'repartition_par_sous_type', 'repartition_par_statut']:
        if key in reps:
            titre = key.replace('repartition_par_', 'Répartition par ').replace('_', ' ').title()
            pdf.set_font(base_font, 'B', 10)
            pdf.cell(0, 6, titre, ln=1)
            pdf.set_font(base_font, '', 8)
            df_source = reps[key]
            # Pour province avec Region colonne explicite
            if key == 'repartition_par_province' and 'Region' in df_source.columns:
                df_rep = df_source.copy()
            else:
                df_rep = df_source.rename(columns=COL_LABELS)
            df_rep = df_rep.head(15)
            cols = df_rep.columns.tolist()
            col_widths = []
            # Largeur dynamique: catégorie 60 puis repartir  (max 3 colonnes esperées)
            first_col_width = 40 if key == 'repartition_par_province' else 60
            remaining = 190 - first_col_width
            if len(cols):
                per = remaining / len(cols)
                col_widths = [per for _ in cols]
            # En-têtes
            pdf.set_fill_color(230, 230, 230)
            label_first = header_first_map.get(key, 'Catégorie')
            if key == 'repartition_par_province':
                # Colonnes: N° | Province | Région (fusionnée) | métriques
                num_col_w = 12
                pdf.cell(num_col_w, 6, 'N°', border=1, fill=True)
                pdf.cell(first_col_width, 6, 'Province', border=1, fill=True)
                pdf.cell(first_col_width, 6, 'Région', border=1, fill=True)
                metric_cols = [c for c in cols if c not in ('Region','Province')]
                width_metrics_total = 190 - first_col_width - first_col_width - num_col_w
                for col in metric_cols:
                    pdf.cell( width_metrics_total / max(1,len(metric_cols)), 6, COL_LABELS.get(col,col), border=1, fill=True)
                pdf.ln(6)
                # Lignes avec fusion logique (répéter région seulement si change)
                last_region = None
                numero = 1
                for _, row in df_rep.iterrows():
                    reg = row.get('Region','')
                    prov = row.get('Province','')
                    # Numéro + Province toujours
                    pdf.cell(num_col_w, 6, str(numero), border=1)
                    pdf.cell(first_col_width, 6, str(prov)[:18], border=1)
                    # Région fusionnée visuellement (cellule vide si répétée)
                    if reg != last_region:
                        pdf.cell(first_col_width, 6, str(reg)[:18], border=1)
                        last_region = reg
                    else:
                        pdf.cell(first_col_width, 6, '', border=1)
                    for col in metric_cols:
                        val = row[col]
                        if isinstance(val, (int,float)) and not isinstance(val,bool):
                            if float(val).is_integer():
                                val_str = f"{int(val):,}".replace(',', ' ')
                            else:
                                val_str = f"{val:,.2f}".replace(',', ' ').replace('.', ',')
                        else:
                            val_str = str(val)
                        pdf.cell( width_metrics_total / max(1,len(metric_cols)), 6, val_str, border=1)
                    pdf.ln(6)
                    numero += 1
                pdf.ln(4)
                continue
            pdf.cell(first_col_width, 6, label_first, border=1, fill=True)
            for w, col in zip(col_widths, cols):
                pdf.cell(w, 6, col, border=1, fill=True)
            pdf.ln(6)
            for idx, row in df_rep.iterrows():
                pdf.cell(first_col_width, 6, str(idx)[:30], border=1)
                for w, col in zip(col_widths, cols):
                    val = row[col]
                    if isinstance(val, (int, float)) and not isinstance(val, bool):
                        val_str = f"{val:,.0f}".replace(',', ' ')
                    else:
                        val_str = str(val)
                    pdf.cell(w, 6, val_str, border=1)
                pdf.ln(6)
            pdf.ln(4)
    # Images
    for img in images_paths:
        pdf.add_page()
        pdf.set_font(base_font, 'B', 11)
        pdf.cell(0, 6, img.stem.replace('_', ' ').title(), ln=1)
        try:
            pdf.image(str(img), w=180)
        except Exception as e:
            pdf.set_font(base_font, '', 9)
            pdf.multi_cell(0, 5, f"Image introuvable: {e}")
    # Page explications & notation
    pdf.add_page()
    pdf.set_font(base_font, 'B', 12)
    pdf.cell(0, 8, _sanitize('4. Notes & Explications'), ln=1)
    pdf.set_font(base_font, '', 9)
    notes = [
        "Notation 1 877 340 (124) : le premier nombre est le montant total (espaces = séparateurs de milliers), le nombre entre parenthèses est le nombre de prestations.",
        "Abréviations des montants sur certaines courbes : 12,3 k = 12 300 ; 1,45 M = 1 450 000 ; 2,10 Md = 2 100 000 000.",
    "Montant total par acte : compare le coût cumulé par type d'acte (consultations, examens, imagerie, pharmacie, etc.) pour identifier les postes majeurs.",
    "Répartition des montants par acte (camembert) : part relative de chaque catégorie d'actes dans la dépense totale.",
    "Top 10 centres (montant / nombre) : centres présentant les montants remboursés les plus élevés et/ou le plus grand nombre de prestations.",
    "Top 10 partenaires : partenaires avec les montants remboursés les plus élevés et/ou le plus grand nombre de prestations.",
        "Répartition par statut de traitement : proportion de prestations acceptées vs autres statuts.",
        "Répartition par région / province : concentration géographique des montants et volumes.",
        "Évolution mensuelle : barres = nombre de prestations, ligne = montant total, annotations = montant abrégé + (nombre).",
        "Évolution trimestrielle : tendance consolidée par trimestre (montant + nombre).",
        "Top 10 adhérents (montant) : adhérents dont le total des prestations est le plus élevé (indicateur de concentration du risque)."
    ]
    for n in notes:
        # simple wrapping
        max_w = pdf.w - pdf.l_margin - pdf.r_margin
        words = n.split()
        line = "- "
        for w in words:
            test = line + w + ' '
            if pdf.get_string_width(test) > max_w:
                pdf.cell(0, 5, _sanitize(line.rstrip()), ln=1)
                line = "  " + w + ' '
            else:
                line = test
        if line.strip():
            pdf.cell(0, 5, _sanitize(line.rstrip()), ln=1)
    pdf.output(OUTPUT_PDF)
    print(f"Export PDF: {OUTPUT_PDF}")

# ------------------ Main ------------------ #

def main():
    df = charger(SOURCE_FILE)
    global_indic = compute_global(df)
    reps = repartitions(df)
    evo, comp_t1 = analyse_temporelle(df)

    # Graphiques principaux
    montant_col = next((c for c in df.columns if c.lower().startswith('montant')), None)
    images_paths = []
    if 'repartition_par_acte' in reps:
        acte_sum = reps['repartition_par_acte']['sum']
        acte_count = reps['repartition_par_acte']['count'] if 'count' in reps['repartition_par_acte'].columns else None
        p1 = plot_bar(acte_sum, 'Montant total par acte', 'montant_par_acte.png', counts_series=acte_count)
        if p1: images_paths.append(p1)
        p1b = plot_pie_group_small(reps['repartition_par_acte']['sum'], 'Répartition des montants par acte', 'montant_par_acte_pie.png')
        if p1b: images_paths.append(p1b)
    if 'repartition_par_centre' in reps:
        centre_sum = reps['repartition_par_centre']['sum'].head(10)
        centre_count = reps['repartition_par_centre']['count'] if 'count' in reps['repartition_par_centre'].columns else None
        if centre_count is not None:
            centre_count = centre_count.head(10)
        p2 = plot_bar(centre_sum, 'Top 10 centres (montant)', 'top10_centres.png', horizontal=True, counts_series=centre_count)
        if p2: images_paths.append(p2)
        p2b = plot_bar(reps['repartition_par_centre']['count'].head(10), 'Top 10 centres (nombre prestations)', 'top10_centres_nbp.png', horizontal=True)
        if p2b: images_paths.append(p2b)
    if 'repartition_par_partenaire' in reps:
        part_sum = reps['repartition_par_partenaire']['sum'].head(10)
        part_count = reps['repartition_par_partenaire']['count'] if 'count' in reps['repartition_par_partenaire'].columns else None
        if part_count is not None:
            part_count = part_count.head(10)
        pp1 = plot_bar(part_sum, 'Top 10 partenaires (montant)', 'top10_partenaires.png', horizontal=True, counts_series=part_count)
        if pp1: images_paths.append(pp1)
        pp2 = plot_bar(reps['repartition_par_partenaire']['count'].head(10), 'Top 10 partenaires (nombre prestations)', 'top10_partenaires_nbp.png', horizontal=True)
        if pp2: images_paths.append(pp2)
    # Top 10 adhérents (montant total)
    montant_col = next((c for c in df.columns if c.lower().startswith('montant')), None)
    adherent_col = next((c for c in df.columns if 'adherent_code' in c.lower()), None)
    if montant_col and adherent_col and adherent_col in df.columns:
        grp_ad = df.groupby(adherent_col)
        top_adherents = grp_ad[montant_col].sum().sort_values(ascending=False).head(10)
        counts_adherents = grp_ad[montant_col].count().reindex(top_adherents.index)
        pa = plot_bar(top_adherents, 'Top 10 adhérents (montant)', 'top10_adherents_montant.png', horizontal=True, counts_series=counts_adherents)
        if pa: images_paths.append(pa)
    if 'repartition_par_statut' in reps:
        ps = plot_pie_group_small(reps['repartition_par_statut']['count'], 'Répartition des prestations par statut', 'repartition_statut_pie.png', threshold=0.01)
        if ps: images_paths.append(ps)
    # Graphiques région & province
    reg_path = plot_region_distribution(reps)
    if reg_path: images_paths.append(reg_path)
    prov_path = plot_province_distribution(reps)
    if prov_path: images_paths.append(prov_path)
    if not evo.empty and montant_col:
        # Préparation données (sans moyenne mobile)
        evo_ma = evo.copy()
        x_vals = list(range(len(evo_ma.index)))
        fig, ax1 = plt.subplots(figsize=(13,5))
        # Barres pour nb prestations
        ax1.bar(x_vals, evo_ma['nb_prestations'], color='#b0c4de', alpha=0.65, label='Nombre de prestations')
        ax1.set_ylabel('Nombre de prestations')
        ax1.set_xlabel('Mois')
        ax2 = ax1.twinx()
        ax2.plot(x_vals, evo_ma['montant_total'], marker='o', color='#d35400', label='Montant total')
        ax2.set_ylabel('Montant (FR)')
        ax2.set_title('Évolution mensuelle – Montants & Prestations')
        # Légendes combinées
        lines, labels = [], []
        for ax in [ax1, ax2]:
            l, lab = ax.get_legend_handles_labels()
            lines.extend(l); labels.extend(lab)
        ax2.legend(lines, labels, loc='upper left')
        # Annotations abrégées montants
        amt_texts = [f"{format_abbrev(v)}\n({int(n)})" for v, n in zip(evo_ma['montant_total'], evo_ma['nb_prestations'])]
        annotate_line_no_overlap(ax2, x_vals, list(evo_ma['montant_total']), amt_texts, min_gap_frac=0.025, base_x_offset=0.4)
        ax1.set_xticks(x_vals)
        ax1.set_xticklabels(evo_ma.index, rotation=45, ha='right')
        ax2.grid(axis='y', linestyle=':', alpha=0.9, linewidth=0.1)
        fig.tight_layout()
        # Bordure carrée figure
        rect = Rectangle((0.005,0.005),0.99,0.99, transform=fig.transFigure, fill=False, lw=1.4, edgecolor='#333')
        fig.patches.append(rect)
        evol_path = FIG_DIR / 'evolution_mensuelle.png'
        fig.savefig(evol_path)
        plt.close(fig)
        images_paths.append(evol_path)
        # Évolution trimestrielle
        date_col = next((c for c in df.columns if c.lower() == 'date'), None)
        if date_col:
            temp_q = df[[date_col, montant_col]].copy()
            temp_q['trimestre'] = temp_q[date_col].dt.to_period('Q').astype(str)
            q_evo = temp_q.groupby('trimestre')[montant_col].agg(['sum','count']).rename(columns={'sum':'montant_total','count':'nb_prestations'})
            if not q_evo.empty:
                fig2, ax2 = plt.subplots(figsize=(9,5))
                sns.lineplot(x=q_evo.index, y=q_evo['montant_total'], marker='o', ax=ax2)
                ax2.set_title('Évolution trimestrielle des montants')
                ax2.set_xlabel('Trimestre')
                ax2.set_ylabel('Montant total')
                for x, (y, n) in zip(q_evo.index, zip(q_evo['montant_total'], q_evo['nb_prestations'])):
                    txt = f"{int(y):,} ({int(n)})".replace(',', ' ')
                    ax2.annotate(txt, (x, y), textcoords='offset points', xytext=(0,7), ha='center', fontsize=8)
                plt.xticks(rotation=0)
                plt.tight_layout()
                # Bordure carrée figure
                rect2 = Rectangle((0.005,0.005),0.99,0.99, transform=fig2.transFigure, fill=False, lw=1.4, edgecolor='#333')
                fig2.patches.append(rect2)
                evol_q_path = FIG_DIR / 'evolution_trimestrielle.png'
                fig2.savefig(evol_q_path)
                plt.close(fig2)
                images_paths.append(evol_q_path)

    export_excel(global_indic, reps, evo, comp_t1, df)
    generer_rapport_html(global_indic, reps, evo, comp_t1, images_paths)
    export_pdf(global_indic, reps, images_paths, evo, comp_t1)

    print('--- Résumé ---')
    for k, v in global_indic.items():
        print(f"{INDIC_LABELS.get(k,k)}: {v}")
    print(f"Export Excel: {OUTPUT_EXCEL}")
    print(f"Graphiques dans: {FIG_DIR}/")
    print("Rapport HTML: rapport_prestations.html")
    if 'top10_centres' in reps:
        print('Top centres (montant):')
        print(reps['top10_centres'].head(5))

def _img_tag_from_path(path: Path) -> str:
    try:
        with open(path, 'rb') as f:
            b64 = base64.b64encode(f.read()).decode('utf-8')
        return f'<img src="data:image/png;base64,{b64}" alt="{path.stem}" style="max-width:100%;height:auto;" />'
    except Exception as e:
        return f'<p>Image {path} indisponible ({e})</p>'

def generer_rapport_html(global_indic: dict, reps: dict, evo: pd.DataFrame, comp_t1: pd.DataFrame, images_paths):
    def fmt(v):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            if pd.isna(v):
                return ''
            # entier ?
            if float(v).is_integer():
                return f"{int(v):,}".replace(',', ' ')
            # décimal 2 chiffres, séparateur milliers espace, virgule française
            s = f"{v:,.2f}"  # 1,234.56
            s = s.replace(',', ' ').replace('.', ',')
            return s
        return str(v)
    # Exclure des indicateurs globaux : montants payés / non payés (ils vont dans Traitement des données)
    exclude_keys = {'montant_paye_total','montant_non_paye','pourcentage_paye_pct'}
    rows_indic = '\n'.join(
        f'<tr><td>{INDIC_LABELS.get(k,k)}</td><td>{fmt(v if not isinstance(v, pd.Timestamp) else v.strftime("%d/%m/%Y"))}</td></tr>'
        for k, v in global_indic.items() if k not in exclude_keys)
    # Ajout synthèse doublons (compteurs) dans indicateurs globaux
    if DUPLICATES_INFO:
        synth_dups = [
            ('Doublons (lignes complètes)', DUPLICATES_INFO.get('nb_doublons_lignes', 0)),
            ("Doublons (identifiant)", DUPLICATES_INFO.get('nb_doublons_identifiant', 0)),
            ("Doublons (clé composite)", DUPLICATES_INFO.get('nb_doublons_cle_composite', 0)),
        ]
        for label, val in synth_dups:
            rows_indic += f"<tr><td>{label}</td><td>{val}</td></tr>"
    sections = []
    # Descriptions des tableaux (inline)
    table_desc = {
        'repartition_par_acte': (
            "Ce tableau dresse un panorama des actes médicaux en rapprochant montants engagés et fréquence. "
            "Il permet d'identifier rapidement les catégories qui structurent la dépense et d'alimenter des pistes d'action : prévention ciblée, renégociation ou contrôle approfondi."
        ),
        'repartition_par_centre': (
            "Nous observons ici la contribution financière et opérationnelle de chaque centre partenaire. "
            "La lecture croisée montants / nombre de prestations aide à distinguer les structures à forte intensité financière de celles à forte intensité d'actes, deux profils justifiant des approches de pilotage différentes."
        ),
        'repartition_par_region': (
            "La vue régionale met en évidence la distribution géographique des remboursements. "
            "Elle sert de premier niveau pour repérer des déséquilibres territoriaux, une concentration inattendue ou des zones nécessitant un renforcement de l'offre."
        ),
        'repartition_par_province': (
            "Le détail par province affine l'analyse territoriale. "
            "En descendant d'un niveau, on valide si les écarts régionaux proviennent d'un noyau restreint de provinces ou d'un phénomène généralisé."
        ),
        'repartition_par_type': (
            "Les grandes familles d'actes sont ici agrégées pour offrir une lecture synthétique. "
            "Cette hiérarchisation facilite la priorisation des segments à surveiller avant d'entrer dans le détail des sous‑types."
        ),
        'repartition_par_sous_type': (
            "Les sous‑types précisent les dynamiques internes à chaque famille. "
            "On identifie les niches coûteuses ou en croissance qui pourraient nécessiter des protocoles, un encadrement tarifaire ou une action de sensibilisation."
        ),
        'repartition_par_statut': (
            "La distribution des statuts renseigne sur la qualité du processus de traitement. "
            "Un niveau élevé de rejets ou d'états intermédiaires orienterait vers des améliorations de saisie, formation ou contrôles."
        ),
        'evolution_mensuelle_tableau': (
            "Le détail mois par mois assure traçabilité et cohérence avec les graphiques. "
            "Il sert aussi de base à toute extrapolation budgétaire ou simulation prospective simple."
        ),
        'evolution_trimestrielle_tableau': (
            "La consolidation trimestrielle atténue les fluctuations accidentelles et rend lisibles les inflexions structurelles. "
            "Elle est adaptée aux présentations de synthèse destinées à la gouvernance."
        )
    }
    # Tableaux principaux
    header_first_map = {
        'repartition_par_acte': 'Acte',
        'repartition_par_centre': 'Centre de santé',
        'repartition_par_region': 'Région',
        'repartition_par_province': 'Province',  # Province + Région handled séparément
        'repartition_par_type': 'Type',
        'repartition_par_sous_type': 'Sous-type',
        'repartition_par_statut': 'Statut'
    }
    for key in ['repartition_par_acte','repartition_par_centre','repartition_par_region','repartition_par_province','repartition_par_type','repartition_par_sous_type','repartition_par_statut']:
        if key in reps:
            tbl = reps[key].copy()
            tbl_fmt = tbl.copy()
            for col in tbl_fmt.columns:
                if col in {'sum','count','mutualistes_distincts'}:
                    tbl_fmt[col] = tbl_fmt[col].apply(lambda x: fmt(x))
                elif col == 'mean':
                    tbl_fmt[col] = tbl_fmt[col].apply(lambda x: fmt(x))
            tbl_fmt.rename(columns=COL_LABELS, inplace=True)
            original_len = len(tbl_fmt)
            if key == 'repartition_par_province' and 'Region' in tbl.columns:
                # Construire HTML avec fusion région (rowspan) + formatage numérique
                dfp = tbl.copy()
                metric_cols = [c for c in dfp.columns if c not in ['Region','Province']]
                for mc in metric_cols:
                    dfp[mc] = dfp[mc].apply(lambda x: fmt(x))
                rows_html = []
                last_region = None
                region_counts = dfp['Region'].value_counts()
                region_rowspans = region_counts.to_dict()
                num = 1
                for _, r in dfp.iterrows():
                    region = r['Region']
                    province = r['Province']
                    row_cells = [f"<td>{num}</td>", f"<td>{province}</td>"]
                    if region != last_region:
                        rowspan = region_rowspans.get(region,1)
                        row_cells.append(f"<td rowspan='{rowspan}' style='font-weight:600;background:#f9f9f9;'>{region}</td>")
                        last_region = region
                    for met_col in metric_cols:
                        row_cells.append(f"<td>{r[met_col]}</td>")
                    rows_html.append('<tr>' + ''.join(row_cells) + '</tr>')
                    num += 1
                header = '<tr><th>Numéro</th><th>Province</th><th>Région</th>' + ''.join(f'<th>{COL_LABELS.get(c,c)}</th>' for c in metric_cols) + '</tr>'
                html_table_inner = f"<table class='tbl'><thead>{header}</thead><tbody>{''.join(rows_html)}</tbody></table>"
                html_table = f"<div class='scroll-box'>{html_table_inner}</div>" if original_len>20 else html_table_inner
                titre_tbl = key.replace('repartition_par_', 'Répartition par ').replace('_',' ').title()
                desc = table_desc.get(key, '')
                intro = f"<p class='expl'>{desc}</p>" if desc else ''
                sections.append(f"<h3>{titre_tbl}</h3>{intro}{html_table}")
                continue
            else:
                # Table simple (toutes lignes) avec scroll si >20 et en-tête personnalisé
                # Remonter l'index en colonne pour contrôler l'en-tête
                if isinstance(tbl_fmt, pd.DataFrame):
                    tbl_fmt = tbl_fmt.reset_index()
                    # Ajout colonne Numéro sauf statut
                    if key != 'repartition_par_statut':
                        tbl_fmt.insert(0, 'Numéro', range(1, len(tbl_fmt) + 1))
                    # Renommer colonne catégorie (juste après Numéro si présente)
                    cat_col_idx = 1 if 'Numéro' in tbl_fmt.columns else 0
                    cat_col = tbl_fmt.columns[cat_col_idx]
                    tbl_fmt.rename(columns={cat_col: header_first_map.get(key, cat_col)}, inplace=True)
                # Construire manuellement la table HTML pour éviter double ligne d'en-tête
                headers = ''.join(f'<th>{h}</th>' for h in tbl_fmt.columns)
                body_rows = []
                for _, r in tbl_fmt.iterrows():
                    cells = ''.join(f'<td>{r[c]}</td>' for c in tbl_fmt.columns)
                    body_rows.append(f'<tr>{cells}</tr>')
                html_table_inner = f"<table class='tbl'><thead><tr>{headers}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"
                html_table = f"<div class='scroll-box'>{html_table_inner}</div>" if original_len>20 else html_table_inner
                if key == 'repartition_par_statut':
                    titre = 'Répartition par statut de traitement'
                else:
                    titre = key.replace('repartition_par_', 'Répartition par ').replace('_', ' ').title()
                desc = table_desc.get(key, '')
                intro = f"<p class='expl'>{desc}</p>" if desc else ''
                sections.append(f'<h3>{titre}</h3>{intro}{html_table}')

    # Tableau d'évolution mensuelle (en plus du graphique)
    if isinstance(evo, pd.DataFrame) and not evo.empty:
        evo_tbl = evo.copy()
        # Normaliser index en colonne Mois
        if evo_tbl.index.name is None:
            evo_tbl = evo_tbl.reset_index().rename(columns={'index': 'Mois'})
        else:
            evo_tbl = evo_tbl.reset_index().rename(columns={evo_tbl.index.name: 'Mois'})
        # Reformater 'Mois' en libellés français complets (ex: janvier 2024)
        try:
            # evo a des clés type '2024-01'; parser
            def _fmt_month(m):
                try:
                    dt = pd.Period(m, freq='M').to_timestamp()
                except Exception:
                    return m
                mois_fr = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'][dt.month-1]
                return f"{mois_fr} {dt.year}"
            evo_tbl['Mois'] = evo_tbl['Mois'].apply(_fmt_month)
        except Exception:
            pass
        # Formater colonnes numériques
        for col in evo_tbl.columns:
            if col.lower().startswith('montant') or col in {'montant_total','sum'}:
                evo_tbl[col] = evo_tbl[col].apply(lambda x: fmt(x))
            elif col in {'nb_prestations','count'}:
                evo_tbl[col] = evo_tbl[col].apply(lambda x: fmt(x))
        # Ajouter Numéro
        evo_tbl.insert(0, 'Numéro', range(1, len(evo_tbl) + 1))
        # Renommer colonnes pour affichage pro
        evo_tbl.rename(columns={'montant_total': 'Montant total', 'nb_prestations': 'Nombre de prestations'}, inplace=True)
        headers = ''.join(f"<th>{h}</th>" for h in evo_tbl.columns)
        body_rows = []
        for _, r in evo_tbl.iterrows():
            cells = ''.join(f"<td>{r[c]}</td>" for c in evo_tbl.columns)
            body_rows.append(f"<tr>{cells}</tr>")
        html_table_inner = f"<table class='tbl'><thead><tr>{headers}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"
        html_table = f"<div class='scroll-box'>{html_table_inner}</div>" if len(evo_tbl) > 20 else html_table_inner
    desc = table_desc.get('evolution_mensuelle_tableau','')
    intro = f"<p class='expl'>{desc}</p>" if desc else ''
    sections.append(f"<h3>Évolution mensuelle (tableau)</h3>{intro}{html_table}")

    # Tableau d'évolution trimestrielle
    if isinstance(evo, pd.DataFrame) and not evo.empty:
        # Reconstituer trimestre à partir de la période mensuelle
        try:
            evo_q = evo.copy()
            evo_q = evo_q.reset_index().rename(columns={evo_q.index.name or 'index': 'Mois'})
            # Convertir en période mensuelle puis en trimestre
            def _to_q(m):
                try:
                    p = pd.Period(m, freq='M')
                    q = p.asfreq('Q')
                    # Format Q1 2024 -> T1 2024
                    return f"T{q.quarter} {q.year}"
                except Exception:
                    return m
            evo_q['Trimestre'] = evo_q['Mois'].apply(_to_q)
            q_agg = evo_q.groupby('Trimestre')[['montant_total','nb_prestations']].sum().reset_index()
            # Trier par année/trimestre chronologiquement
            def _sort_key(t):
                try:
                    parts = t.split()
                    q = int(parts[0][1:])
                    y = int(parts[1])
                    return (y, q)
                except Exception:
                    return (9999, 99)
            q_agg = q_agg.sort_values(key=lambda s: s.apply(_sort_key))
            # Formater
            q_agg.insert(0, 'Numéro', range(1, len(q_agg)+1))
            for col in ['montant_total','nb_prestations']:
                if col in q_agg.columns:
                    q_agg[col] = q_agg[col].apply(lambda x: fmt(x))
            q_agg.rename(columns={'montant_total':'Montant total','nb_prestations':'Nombre de prestations'}, inplace=True)
            headers_q = ''.join(f"<th>{h}</th>" for h in q_agg.columns)
            body_rows_q = []
            for _, r in q_agg.iterrows():
                cells = ''.join(f"<td>{r[c]}</td>" for c in q_agg.columns)
                body_rows_q.append(f"<tr>{cells}</tr>")
            html_table_inner_q = f"<table class='tbl'><thead><tr>{headers_q}</tr></thead><tbody>{''.join(body_rows_q)}</tbody></table>"
            html_table_q = f"<div class='scroll-box'>{html_table_inner_q}</div>" if len(q_agg) > 20 else html_table_inner_q
            desc_q = table_desc.get('evolution_trimestrielle_tableau','')
            intro_q = f"<p class='expl'>{desc_q}</p>" if desc_q else ''
            sections.append(f"<h3>Évolution trimestrielle (tableau)</h3>{intro_q}{html_table_q}")
        except Exception:
            pass
    # Construction des blocs graphiques avec description inline
    graph_desc = {
        'montant_par_acte': (
            "Montant total par acte",
            "Ce graphique hiérarchise les actes selon leur poids financier. Il met immédiatement en relief les leviers potentiels où une action ciblée (prévention, protocole, renégociation) aurait l'impact le plus sensible."
        ),
        'montant_par_acte_pie': (
            "Répartition des montants par acte",
            "La part relative de chaque acte confirme ou nuance la concentration observée. Une forte polarisation incite à analyser les déterminants cliniques ou organisationnels de ces catégories dominantes."
        ),
        'top10_centres': (
            "Top 10 centres – montants",
            "La concentration financière par centre éclaire les priorités de dialogue et d'audit. Les structures en tête reflètent soit un volume important d'actes lourds soit une spécialisation coûteuse."
        ),
        'top10_centres_nbp': (
            "Top 10 centres – nombre de prestations",
            "La dynamique d'activité (fréquence) peut différer de la dynamique financière. Identifier ce décalage aide à ajuster l'allocation des contrôles ou des conventions."
        ),
        'top10_partenaires': (
            "Top 10 partenaires – montants",
            "Les partenaires externes majeurs représentent des points de dépendance opérationnelle. Leur suivi permet d'anticiper tout risque de rupture ou de dérive tarifaire."
        ),
        'top10_partenaires_nbp': (
            "Top 10 partenaires – nombre de prestations",
            "Un partenaire très sollicité en nombre d'actes mais modeste en montant peut être un pivot logistique (soins courants) à préserver en termes de qualité et de fluidité."
        ),
        'top10_adherents_montant': (
            "Top 10 adhérents – montants",
            "La concentration sur quelques adhérents peut révéler des pathologies chroniques ou des parcours complexes. Elle suggère des actions de prévention ou de coordination renforcée."
        ),
        'repartition_statut_pie': (
            "Répartition par statut de traitement",
            "La ventilation des statuts sert de baromètre du processus administratif. Un niveau anormal de rejets ou d'états temporaires déclenche une vérification de la chaîne de saisie et validation."
        ),
        'evolution_mensuelle': (
            "Évolution mensuelle – graphique",
            "La combinaison barres (nombre) et courbe (montant) met en évidence divergences : hausse de coût sans hausse d'actes ou inversement. Ces inflexions guident les analyses causales."
        ),
        'evolution_trimestrielle': (
            "Évolution trimestrielle – graphique",
            "La vision lissée confirme si les variations récentes constituent un signal durable ou un bruit ponctuel."
        ),
        'repartition_region_montant_nombre': (
            "Répartition géographique – régions",
            "Comparer activité et charge financière par région révèle des profils distincts (intensité de recours vs coût moyen)."
        ),
        'repartition_province_montant_nombre': (
            "Répartition géographique – provinces",
            "Le niveau provincial précise les foyers locaux à suivre, notamment quand une région est hétérogène."
        ),
        'pareto_actes': (
            "Courbe de Pareto des actes",
            "La courbe cumulative illustre la part de dépense expliquée par un noyau restreint d'actes. Elle justifie une segmentation prioritaire pour optimiser l'effort de gestion."
        ),
        'scatter_types': (
            "Dispersion actes (montant moyen vs fréquence)",
            "La matrice fréquence / montant moyen positionne les actes : ceux combinant coût unitaire élevé et fréquence notable représentent la zone critique d'optimisation."
        )
    }
    graph_blocks = []
    for p in images_paths:
        stem = p.stem
        titre, desc = graph_desc.get(stem, (stem.replace('_',' ').title(), "Graphique de suivi."))
        img_tag = _img_tag_from_path(p)
        graph_blocks.append(f"<div class='visu'><h3>{titre}</h3><p class='expl'>{desc}</p>{img_tag}</div>")
    imgs_html = '\n'.join(graph_blocks)
    objectifs_html = """
    <section>
        <h2>Objectif de l’analyse</h2>
        <ul>
            <li>Évaluer les dépenses de santé prises en charge par la mutuelle.</li>
            <li>Comparer la répartition par actes, centres, partenaires et zones géographiques.</li>
            <li>Mesurer la performance du remboursement (taux d’acceptation, paiements effectués).</li>
            <li>Identifier les tendances pour anticiper la charge financière future.</li>
        </ul>
    </section>
    """
    explications_html = ""  # remplacé par descriptions inline
    # Section Traitement des données
    montant_paye = global_indic.get('montant_paye_total', '')
    montant_non_paye = global_indic.get('montant_non_paye', '')
    pct_paye = global_indic.get('pourcentage_paye_pct', '')
    nb_dates_corr = 0
    if 'DATES_CORRIGEES_DF' in globals() and DATES_CORRIGEES_DF is not None:
        nb_dates_corr = len(DATES_CORRIGEES_DF)
    traitement_rows = []
    def _fmt_brut(v):
        if isinstance(v,(int,float)) and not isinstance(v,bool):
            if float(v).is_integer():
                return f"{int(v):,}".replace(',', ' ')
            return f"{v:,.2f}".replace(',', ' ').replace('.', ',')
        return v
    # (Montant payé / non payé retirés sur demande)
    if nb_dates_corr:
        traitement_rows.append(f"<tr><td>Dates corrigées (août–déc. 2025 =&gt; 2024)</td><td>{nb_dates_corr}</td></tr>")
    if DUPLICATES_INFO:
        traitement_rows.append(f"<tr><td>Lignes doublons supprimées</td><td>{DUPLICATES_INFO.get('nb_doublons_lignes',0)}</td></tr>")
    traitement_html = f"""
    <section>
      <h2>1. Traitement des données</h2>
      <p>Opérations appliquées avant analyse :</p>
      <ul>
        <li>Normalisation des noms de colonnes / centres (majuscules, espaces).</li>
        <li>Conversion des dates, filtrage sur 2024–2025.</li>
        <li>Correction : toutes les dates d'août à décembre 2025 ont leur année remplacée par 2024.</li>
        <li>Nettoyage des montants (espaces / séparateurs) et conversion numérique.</li>
        <li>Renommage de « validite » en « statut » + uniformisation (minuscules).</li>
        <li>Suppression des doublons (lignes identiques) après extraction vers feuilles dédiées.</li>
        <li>Calcul montants payés : somme des montants où statut ∈ {"accepté","accepte"}; non payé = total - payé.</li>
      </ul>
      <table class='tbl'><thead><tr><th>Élément</th><th>Valeur</th></tr></thead><tbody>{''.join(traitement_rows) if traitement_rows else '<tr><td>(aucun)</td><td>-</td></tr>'}</tbody></table>
    <!-- Ligne de formules supprimée sur demande -->
    </section>
    """
    # Logo base64 si dispo
    logo_html = ''
    if LOGO_PATH.exists():
        try:
            with open(LOGO_PATH, 'rb') as lf:
                _b64 = base64.b64encode(lf.read()).decode('utf-8')
            logo_html = f"<img src='data:image/png;base64,{_b64}' alt='Logo' style='height:70px;float:right;margin-left:15px;' />"
        except Exception:
            logo_html = ''
    # Logo pour le footer (version réduite)
    footer_logo_html = ''
    if LOGO_PATH.exists():
        try:
            with open(LOGO_PATH, 'rb') as lf2:
                _b64_footer = base64.b64encode(lf2.read()).decode('utf-8')
            footer_logo_html = f"<img src='data:image/png;base64,{_b64_footer}' alt='Logo' style='height:34px;display:inline-block;' />"
        except Exception:
            footer_logo_html = ''
    html = f"""<!DOCTYPE html><html lang='fr'><head><meta charset='utf-8'><title>Rapport Prestations</title>
<style>
body{{font-family:Arial,Helvetica,sans-serif;margin:25px;}}
h1{{color:#2E4053;}}
table.tbl{{border-collapse:collapse;margin:10px 0;font-size:13px;}}
table.tbl th,table.tbl td{{border:1px solid #ccc;padding:4px 8px;text-align:left;}}
thead th{{background:#f2f2f2;}}
.grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:20px;}}
figure{{margin:0;}}
figcaption{{font-size:12px;text-align:center;margin-top:4px;color:#555;}}
.scroll-box{{max-height:480px;overflow-y:auto;border:1px solid #bbb;padding:4px;background:#fafafa;margin-top:6px;}}
details summary{{cursor:pointer;font-weight:600;margin:6px 0;}}
</style></head><body>
<div style='overflow:auto;'> {logo_html}<h1 style='margin-top:10px;'>Rapport Analytique des Prestations</h1></div>
{objectifs_html}
{traitement_html}
<h2>3. Indicateurs globaux</h2>
<p class='expl'>Tableau de synthèse générale (périmètre, dispersion, volumes et période d'observation) servant de point d'entrée à l'analyse.</p>
<table class='tbl'><thead><tr><th>Indicateur</th><th>Valeur</th></tr></thead><tbody>{rows_indic}</tbody></table>
<h2>4. Graphiques analytiques</h2>
<div class='grid'>{imgs_html}</div>
<h2>5. Tableaux de répartition détaillés</h2>
{''.join(sections)}
<hr style='margin:40px 0 15px;border:none;border-top:1px solid #ccc;'>
<footer style='text-align:center;font-size:13px;color:#555;font-style:italic;'>
    <div style='display:flex;align-items:center;justify-content:center;gap:10px;'>
        {footer_logo_html}
        <span>Mutuelle de la Police Nationale</span>
    </div>
</footer>
</body></html>"""
    with open('rapport_prestations.html','w',encoding='utf-8') as f:
        f.write(html)


if __name__ == '__main__':
    main()
