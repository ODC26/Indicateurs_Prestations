from pathlib import Path
from export_monthly_actes_word import load_source, compute_monthly_counts
import math


def approx_equal(a, b, tol=1e-6):
    try:
        return abs(float(a) - float(b)) <= tol
    except Exception:
        return a == b


def main():
    src = Path('Classeur1.xlsx')
    df, acte_col, date_col = load_source(src)
    p_counts, p_amt, p_ad, p_ay, overall_ad, overall_ay = compute_monthly_counts(df, acte_col, date_col, year=2025)

    months = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet']

    issues = []

    def check_pivot(pivot, name, float_vals=False):
        for idx, row in pivot.iterrows():
            s = 0.0 if float_vals else 0
            for m in months:
                v = row.get(m, 0)
                s += float(v) if float_vals else int(v)
            total = row.get('Total', 0)
            if float_vals:
                if not approx_equal(s, total):
                    issues.append(f"{name} mismatch for '{idx}': sum months={s} total={total}")
            else:
                if int(s) != int(total):
                    issues.append(f"{name} mismatch for '{idx}': sum months={s} total={total}")

    # Check each pivot
    check_pivot(p_counts, 'Prestations', float_vals=False)
    check_pivot(p_amt, 'Montants', float_vals=True)
    check_pivot(p_ad, 'Adhérents', float_vals=False)
    check_pivot(p_ay, 'Ayants droit', float_vals=False)

    # Check overall totals consistency
    tot_months_counts = sum(p_counts[m].sum() for m in months)
    tot_total_counts = p_counts['Total'].sum() if 'Total' in p_counts.columns else 0
    if int(tot_months_counts) != int(tot_total_counts):
        issues.append(f"Overall counts mismatch: sum(per-month totals)={tot_months_counts} vs sum(Total)={tot_total_counts}")

    tot_months_amt = sum(p_amt[m].sum() for m in months)
    tot_total_amt = p_amt['Total'].sum() if 'Total' in p_amt.columns else 0.0
    if not approx_equal(tot_months_amt, tot_total_amt):
        issues.append(f"Overall montants mismatch: sum(per-month totals)={tot_months_amt} vs sum(Total)={tot_total_amt}")

    if issues:
        print('VALIDATION FAILED:')
        for it in issues:
            print('-', it)
    else:
        print('VALIDATION PASSED: Tous les Totals correspondent à la somme des mois pour chaque acte et au global.')


if __name__ == '__main__':
    main()
