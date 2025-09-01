import pandas as pd
from analyse_prestations_full import charger, repartitions

# Charger les données avec le bon fichier
df = charger('Classeur1.xlsx')
print('Données chargées:', len(df), 'lignes')

# Calculer les répartitions
reps = repartitions(df)

# Afficher le tableau des bénéficiaires
if 'repartition_par_beneficiaire' in reps:
    print('\nTableau des bénéficiaires COMPLET:')
    beneficiaires_df = reps['repartition_par_beneficiaire']
    print(beneficiaires_df)
    
    print('\nDétail par ligne:')
    for idx, row in beneficiaires_df.iterrows():
        print(f"'{idx}': sum={row['sum']}, count={row['count']}, mean={row['mean']:.2f}, mutualistes={row['mutualistes_distincts']}")
    
    # Test du calcul manuel
    mutualistes_total = 0
    if 'mutualistes_distincts' in beneficiaires_df.columns:
        for idx, row in beneficiaires_df.iterrows():
            if isinstance(idx, str) and (idx.lower().startswith('adherent') or idx.lower().startswith('adhérent') or 'ayant' in idx.lower()):
                print(f"  -> Ligne comptée: '{idx}' avec {row['mutualistes_distincts']} mutualistes")
                mutualistes_total += int(row['mutualistes_distincts'])
    
    print(f'\nTotal calculé: {mutualistes_total}')
    print(f'Calcul attendu: 1229 + 102 = 1331')
else:
    print('Pas de répartition par bénéficiaire trouvée')
