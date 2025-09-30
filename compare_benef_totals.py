from pathlib import Path
from export_monthly_actes_word import load_source, compute_monthly_counts
import pandas as pd
import calendar


def compute_overall_benef_counts(df: pd.DataFrame, date_col: str, adherent_col: str, beneficiaire_col: str, year: int = 2025):
    # returns dict month->(ad, ay, total)
    res = {}
    for month in range(1, 8):
        mdata = df[pd.notna(df[date_col]) & (df[date_col].dt.year == year) & (df[date_col].dt.month == month)]
        ad = 0
        ay = 0
        if beneficiaire_col and adherent_col:
            beneficiaire_stats = mdata.groupby(beneficiaire_col).agg({adherent_col: 'nunique'})
            idx = beneficiaire_stats.index.astype(str).str.lower()
            adherents_mask = idx.str.startswith(('adhérent', 'adherent'))
            ayants_mask = idx.str.contains('ayant', regex=False)
            if adherents_mask.any():
                ad = int(beneficiaire_stats.loc[adherents_mask, adherent_col].sum())
            if ayants_mask.any():
                ay = int(beneficiaire_stats.loc[ayants_mask, adherent_col].sum())
        else:
            # fallback: count unique adherent codes if present, else 0
            if adherent_col:
                ad = int(mdata[adherent_col].nunique())
        res[month] = (ad, ay, ad + ay)
    return res


def main():
    src = Path('Classeur1.xlsx')
    df, acte_col, date_col = load_source(src)

    # detect columns
    adherent_col = next((c for c in df.columns if 'adherent_code' in str(c).lower()), None)
    beneficiaire_col = next((c for c in df.columns if str(c).lower() == 'beneficiaire'), None)

    p_counts, p_amt, p_ad, p_ay, overall_ad, overall_ay = compute_monthly_counts(df, acte_col, date_col, year=2025)

    overall = compute_overall_benef_counts(df, date_col, adherent_col, beneficiaire_col, year=2025)

    months = {1: 'Janvier',2:'Février',3:'Mars',4:'Avril',5:'Mai',6:'Juin',7:'Juillet'}

    print('Month | Overall_Ad | Overall_Ay | Overall_Total || PerAct_Ad (sum) | PerAct_Ay (sum) | PerAct_Total | Diff_Total')
    print('-'*110)
    for m in range(1,8):
        mon = months[m]
        o_ad, o_ay, o_tot = overall[m]
        per_ad = int(p_ad[mon].sum()) if mon in p_ad.columns else 0
        per_ay = int(p_ay[mon].sum()) if mon in p_ay.columns else 0
        per_tot = per_ad + per_ay
        diff = per_tot - o_tot
        # also show overall totals computed by compute_monthly_counts
        overall_month_tot = int(overall_ad.get(mon, 0) + overall_ay.get(mon, 0))
        print(f"{mon:7} | {o_ad:10} | {o_ay:11} | {o_tot:13} || {per_ad:16} | {per_ay:16} | {per_tot:12} | {diff:9} || Ensemble_calc: {overall_month_tot}")

    # Summary explanation
    print('\nExplanation:')
    print('- If PerAct_Total > Overall_Total, it means beneficiaries are being counted per-act and summed across acts (same person with multiple acts counted multiple times).')
    print('- If PerAct_Total == Overall_Total, both approaches align (no double-counting across acts).')


if __name__ == '__main__':
    main()
