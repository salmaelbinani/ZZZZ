
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import numpy as np

def integrate_files(separated_file, circuit_file):
    # Chargement des DataFrames depuis les fichiers fournis
    df_mmsta = pd.read_excel(separated_file, sheet_name='FIL SIMPLE')
    df_mmsta1 = pd.read_excel(separated_file, sheet_name='double')
    df_mmsta2 = pd.read_excel(separated_file, sheet_name='twist')
    df_mmsta3 = pd.read_excel(separated_file, sheet_name='SQUIB')
    df_mmsta4 = pd.read_excel(separated_file, sheet_name='GW')
    df_mmsta5 = pd.read_excel(separated_file, sheet_name='joint')
    df_mmsta6 = pd.read_excel(separated_file, sheet_name='super group')
    df_maxwire = pd.read_excel(circuit_file, sheet_name='Report(Draft)')

    # Insertion de nouvelles colonnes après 'Wire Internal Name'
    insert_position = df_maxwire.columns.get_loc('Wire Internal Name') + 1
    df_maxwire.insert(insert_position, 'TYPE', '')
    df_maxwire.insert(insert_position + 1, 'SN ADD', '')
    df_maxwire.insert(insert_position + 2, 'SN P2', '')
    df_maxwire.insert(insert_position + 3, 'Super group', '')

    # Nettoyage et transformation des données
    df_maxwire['Wire Internal Name'] = df_maxwire['Wire Internal Name'].astype(str).str.replace('W', '')
    df_maxwire['Wire Cross-Section'] = df_maxwire['Wire Cross-Section'].apply(lambda x: str(x).replace('.0', '') if str(x).endswith('.0') else str(x))
    df_maxwire['From Seal by Terminal'] = df_maxwire['From Seal by Terminal'].astype(str).str.replace('.0', '')
    df_maxwire['To Seal by Terminal'] = df_maxwire['To Seal by Terminal'].astype(str).str.replace('.0', '')
    df_maxwire['Wire Part Number'] = df_maxwire['Wire Part Number'].astype(str).str.replace('180', '')

    # Création de la colonne 'salma'
    df_maxwire['salma'] = (
        'Circuit ' +
        df_maxwire['Wire Internal Name'].astype(str) + ' ' +
        df_maxwire['Wire Kind'].astype(str) + ' ' +
        df_maxwire['Wire Cross-Section'].astype(str) + ' ' +
        df_maxwire['Wire Color'].astype(str)
    )

   

    # Étape 1: Nettoyage des données dans df_mmsta5 pour ignorer les SN2 vides ou contenant "(blanks)"
    df_mmsta5_clean = df_mmsta5.dropna(subset=['SN2']).copy()  # Supprime les lignes où SN2 est NaN
    df_mmsta5_clean = df_mmsta5_clean[df_mmsta5_clean['SN2'].str.strip() != '(blanks)']  # Supprime les lignes où SN2 est "(blanks)"

    # Étape 2: Fonction pour mapper 'SN ADD' et 'SN P2' en utilisant 'salma' et 'DSN2'
    def map_sn_add_and_sn_p2(salma, df):
        matched_row = df[df['DSN2'] == salma]
        if not matched_row.empty:
            sn_add = matched_row['SN2'].tolist()
            sn_p2 = matched_row['SN1'].tolist()
            return sn_add, sn_p2
        return None, None

    # Fonction pour mettre à jour 'SN ADD' et 'SN P2'
    def update_sn_add_and_sn_p2(row, df_mmsta5_clean):
        salma = row['salma']
        new_sn_add, new_sn_p2 = map_sn_add_and_sn_p2(salma, df_mmsta5_clean)

        # Mise à jour de SN ADD
        if isinstance(row['SN ADD'], list) and any(str(val).strip() for val in row['SN ADD']):
            sn_add_value = row['SN ADD']
        else:
            sn_add_value = new_sn_add if new_sn_add else row['SN ADD']

        # Mise à jour de SN P2
        if isinstance(row['SN P2'], list) and any(str(val).strip() for val in row['SN P2']):
            sn_p2_value = row['SN P2']
        else:
            sn_p2_value = new_sn_p2 if new_sn_p2 else row['SN P2']

        return sn_add_value, sn_p2_value

    # Étape 3: Remplir les colonnes 'SN ADD' et 'SN P2'
    df_maxwire[['SN ADD', 'SN P2']] = df_maxwire.apply(lambda row: update_sn_add_and_sn_p2(row, df_mmsta5_clean), axis=1, result_type='expand')

    # Étape 4: Aplatir les listes des SN ADD et SN P2
    df_maxwire['SN ADD'] = df_maxwire['SN ADD'].apply(lambda x: x if isinstance(x, list) else [x] if pd.notna(x) else [])
    df_maxwire['SN P2'] = df_maxwire['SN P2'].apply(lambda x: x if isinstance(x, list) else [x] if pd.notna(x) else [])

    df_maxwire = df_maxwire.explode(['SN ADD', 'SN P2']).reset_index(drop=True)

    # Étape 5: Remplissage de la colonne 'TYPE'
    df_maxwire['TYPE'] = df_maxwire.apply(
        lambda row: 'Joint' if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '' and (pd.isna(row['TYPE']) or str(row['TYPE']).strip() == '') else row['TYPE'],
        axis=1
    )

    # Étape 6: Mise à jour des colonnes de df_maxwire en fonction des colonnes de df_mmsta5
    columns_to_check = df_mmsta5.columns[df_mmsta5.columns.get_loc('DSN3'):df_mmsta5.columns.get_loc('Total') + 1]
    columns_to_update = df_maxwire.columns[df_maxwire.columns.get_loc('To Eyelet Protection'):df_maxwire.columns.get_loc('salma') + 1]

    for index, row in df_maxwire.iterrows():
        salma = row['salma']
        matched_row = df_mmsta5_clean[df_mmsta5_clean['DSN2'] == salma]

        if not matched_row.empty:
            for col_to_check, col_to_update in zip(columns_to_check, columns_to_update):
                if matched_row.iloc[0][col_to_check] == 1:
                    df_maxwire.loc[index, col_to_update] = 'X'
                else:
                    df_maxwire.loc[index, col_to_update] = np.nan  # Supprimer le "X" s'il ne doit pas être présent



    # Étape 1: Nettoyage des données dans df_mmsta pour ignorer les SN1 vides ou contenant "(blanks)"
    df_mmsta_clean = df_mmsta.dropna(subset=['SN1']).copy()  # Supprime les lignes où SN1 est NaN
    df_mmsta_clean = df_mmsta_clean[df_mmsta_clean['SN1'].str.strip() != '(blanks)']  # Supprime les lignes où SN1 est "(blanks)"

    # Étape 2: Fonction pour mapper 'SN ADD' en utilisant 'salma' et 'DSN1'
    def map_sn_add(salma, df):
        match = df.loc[df['DSN1'] == salma, 'SN1']
        return match.values[0] if not match.empty else None

    # Fonction pour mettre à jour 'SN ADD' uniquement si elle est vide
    def update_sn_add_if_empty(row, df_mmsta_clean):
        # Conserver la valeur existante si elle contient des valeurs non vides
        if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '':
            return row['SN ADD']
        # Sinon, essayer de remplir 'SN ADD' avec les correspondances trouvées
        salma = row['salma']
        new_sn_add = map_sn_add(salma, df_mmsta_clean)
        return new_sn_add if new_sn_add else row['SN ADD']

    # Étape 3: Mettre à jour la colonne 'SN ADD' uniquement si elle est vide
    df_maxwire['SN ADD'] = df_maxwire.apply(lambda row: update_sn_add_if_empty(row, df_mmsta_clean), axis=1)

    # Étape 4: Aplatir la liste des SN ADD (si nécessaire)
    df_maxwire['SN ADD'] = df_maxwire['SN ADD'].apply(lambda x: x if isinstance(x, list) else [x] if pd.notna(x) else [])
    df_maxwire = df_maxwire.explode('SN ADD').reset_index(drop=True)

    # Étape 5: Remplissage de la colonne 'TYPE' avec 'FIL SIMPLE' uniquement si elle est vide et que 'SN ADD' est non vide
    df_maxwire['TYPE'] = df_maxwire.apply(
        lambda row: 'FIL SIMPLE' if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '' and (pd.isna(row['TYPE']) or str(row['TYPE']).strip() == '') else row['TYPE'],
        axis=1
    )

    # Étape 6: Vérifier le nombre de lignes avec 'TYPE' égal à 'FIL SIMPLE'
    num_fil_simple = df_maxwire[df_maxwire['TYPE'] == 'FIL SIMPLE'].shape[0]
    num_separation_rows = df_mmsta_clean.shape[0]
    print(f"Initial count of 'FIL SIMPLE': {num_fil_simple}")
    print(f"Number of rows in df_mmsta_clean: {num_separation_rows}")

    # Étape 7: Identifier les lignes de df_mmsta_clean qui ne sont pas encore intégrées dans df_maxwire
    non_integrated_df = df_mmsta_clean[~df_mmsta_clean['SN1'].isin(df_maxwire['SN ADD'])]

    # Préparer les caractéristiques à vérifier dans 'DS Général'
    characteristics = ['From Seal by Terminal', 'From Terminal', 'To Terminal', 'To Seal by Terminal', 'Wire Part Number', 'Final Wire Length']

    # S'assurer que toutes les colonnes pertinentes sont traitées comme des chaînes et gérer les valeurs NaN
    for char in characteristics:
        df_maxwire.loc[:,char] = df_maxwire[char].astype(str).fillna('')

    non_integrated_df.loc[:,'DS Général'] = non_integrated_df['DS Général'].astype(str).fillna('')

    # Fonction pour vérifier si toutes les caractéristiques pertinentes sont des sous-chaînes de la colonne DS Général
    def check_inclusion(row, ds_general):
        items = [row[char].strip() for char in characteristics if row[char].strip() not in ['', 'nan']]
        return all(item in ds_general for item in items)

    # Fonction pour intégrer les valeurs de 'SN ADD' basées sur les caractéristiques, sans écrasement
    def integrate_sn_add(row, non_integrated_df):
        if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '':
            return row['SN ADD']  # Conserver la valeur existante si elle contient des valeurs non vides
        matched_rows = non_integrated_df[non_integrated_df['DS Général'].apply(lambda x: check_inclusion(row, x))]
        if not matched_rows.empty:
            return matched_rows['SN1'].tolist()  # Retourner toutes les correspondances sous forme de liste
        return []  # Retourner une liste vide si aucune correspondance trouvée

    # Appliquer l'intégration pour les lignes non intégrées uniquement à partir de non_integrated_df
    df_maxwire['SN ADD'] = df_maxwire.apply(lambda row: integrate_sn_add(row, non_integrated_df) if row['TYPE'] != 'FIL SIMPLE' else row['SN ADD'], axis=1)

    # Aplatir à nouveau la liste des SN ADD après intégration
    df_maxwire = df_maxwire.explode('SN ADD').reset_index(drop=True)

    # Étape 8: Mettre à jour la colonne 'TYPE' après l'intégration supplémentaire, si 'TYPE' est encore vide
    df_maxwire['TYPE'] = df_maxwire.apply(
        lambda row: 'FIL SIMPLE' if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '' and (pd.isna(row['TYPE']) or str(row['TYPE']).strip() == '') else row['TYPE'],
        axis=1
    )

    # Vérifier le nombre de lignes avec 'TYPE' égal à 'FIL SIMPLE' après intégration
    num_fil_simple_updated = df_maxwire[df_maxwire['TYPE'] == 'FIL SIMPLE'].shape[0]
    print(f"Updated count of 'FIL SIMPLE': {num_fil_simple_updated}")
    print(f"Number of rows in df_mmsta_clean: {num_separation_rows}")

    # Remplacer les valeurs de 'SN ADD' et 'TYPE' par des chaînes vides lorsque la ligne contient "double"
    print("Replacing 'SN ADD' and 'TYPE' where 'double' is found")
    mask = df_maxwire.apply(lambda row: row.astype(str).str.contains('double', case=False).any(), axis=1)
    df_maxwire.loc[mask, ['SN ADD', 'TYPE']] = ''

  


    # Étape 1: Nettoyage des données dans df_mmsta1 pour ignorer les SN2 vides ou contenant "(blanks)"
    df_mmsta1_clean = df_mmsta1.dropna(subset=['SN2']).copy()  # Supprime les lignes où SN2 est NaN
    df_mmsta1_clean = df_mmsta1_clean[df_mmsta1_clean['SN2'].str.strip() != '(blanks)']  # Supprime les lignes où SN2 est "(blanks)"

    # Étape 2: Remplissage initial de la colonne 'SN ADD' en utilisant une correspondance basée sur 'salma' et 'DSN2'
    def map_sn_add(salma, df):
        matches = df[df['DSN2'] == salma]['SN2']
        return matches.iloc[0] if not matches.empty else ''  # Garder seulement la première correspondance

    # Conserver les valeurs existantes dans 'SN ADD' si elles sont déjà non vides
    def update_sn_add_if_empty(row, df_mmsta1_clean):
        if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '':
            return row['SN ADD']  # Conserver la valeur existante si non vide
        # Sinon, essayer de remplir 'SN ADD' avec la correspondance trouvée
        salma = row['salma']
        if salma in df_mmsta1_clean['DSN2'].values:
            return map_sn_add(salma, df_mmsta1_clean)
        return row.get('SN ADD', '')  # Retourner la valeur existante ou une chaîne vide si aucune correspondance trouvée

    # Étape 3: Remplir la colonne 'SN ADD'
    df_maxwire['SN ADD'] = df_maxwire.apply(lambda row: update_sn_add_if_empty(row, df_mmsta1_clean), axis=1)

    # Étape 4: Mise à jour de la colonne 'TYPE'
    df_maxwire['TYPE'] = df_maxwire.apply(
        lambda row: 'Double' if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '' and (pd.isna(row['TYPE']) or str(row['TYPE']).strip() == '') else row['TYPE'],
        axis=1
    )

    # Étape 5: Conserver les valeurs de 'SN ADD' existantes pour les lignes où 'SN ADD' est déjà non vide
    def integrate_sn_add(row, non_integrated_df1):
        if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '':
            return row['SN ADD']  # Conserver la valeur existante si non vide
        matched_rows = non_integrated_df1[non_integrated_df1['DS Général'].apply(lambda x: check_inclusion(row, x))]
        if not matched_rows.empty:
            return matched_rows['SN2'].iloc[0]
        return ''  # Retourner une chaîne vide si aucune correspondance trouvée

    # Identifier les lignes non intégrées
    non_integrated_df1 = df_mmsta1_clean[~df_mmsta1_clean['SN2'].isin(df_maxwire['SN ADD'])]

    # Préparation des caractéristiques (conversion en chaînes de caractères)
    characteristics = ['From Seal by Terminal', 'From Terminal', 'To Terminal', 'To Seal by Terminal', 'Wire Part Number', 'Final Wire Length']

    for char in characteristics:
        df_maxwire.loc[:,char] = df_maxwire[char].astype(str).fillna('')

    non_integrated_df1.loc[:,'DS Général'] = non_integrated_df1['DS Général'].astype(str).fillna('')

    # Vérification de l'inclusion des caractéristiques
    def check_inclusion(row, ds_general):
        items = [row[char].strip() for char in characteristics if row[char].strip() not in ['', 'nan']]
        return all(item in ds_general for item in items)

    # Mettre à jour les valeurs de 'SN ADD' uniquement si elles sont vides
    df_maxwire['SN ADD'] = df_maxwire.apply(lambda row: integrate_sn_add(row, non_integrated_df1), axis=1)

    # Étape 6: Mettre à jour 'SN P2' là où 'SN ADD' est égale à 'SN2' de df_mmsta1
    for i, row in df_maxwire[df_maxwire['SN ADD'].notna()].iterrows():
        sn_add = row['SN ADD']
        sn_1_value = df_mmsta1.loc[df_mmsta1['SN2'] == sn_add, 'SN1']
        if not sn_1_value.empty:
            df_maxwire.at[i, 'SN P2'] = sn_1_value.values[0]




    # Étape 1: Nettoyage des données dans df_mmsta2 pour ignorer les SN2 vides ou contenant "(blanks)"
    df_mmsta2_clean = df_mmsta2.dropna(subset=['SN2']).copy()  # Supprime les lignes où SN2 est NaN
    df_mmsta2_clean = df_mmsta2_clean[df_mmsta2_clean['SN2'].str.strip() != '(blanks)']  # Supprime les lignes où SN2 est "(blanks)"

    # Étape 2: Remplissage de la colonne 'SN ADD' uniquement si elle est vide
    def map_sn_add(salma, df):
        matches = df[df['DSN2'] == salma]['SN2']
        return matches.iloc[0] if not matches.empty else ''

    # Conserver les valeurs existantes dans 'SN ADD' si elles ne sont pas vides
    def update_sn_add_if_empty(row, df_mmsta2_clean):
        if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '':
            return row['SN ADD']  # Conserver la valeur existante si non vide
        # Sinon, essayer de remplir 'SN ADD' avec la correspondance trouvée
        salma = row['salma']
        if salma in df_mmsta2_clean['DSN2'].values:
            return map_sn_add(salma, df_mmsta2_clean)
        return row.get('SN ADD', '')

    # Étape 3: Remplir la colonne 'SN ADD'
    df_maxwire['SN ADD'] = df_maxwire.apply(lambda row: update_sn_add_if_empty(row, df_mmsta2_clean), axis=1)

    # Étape 4: Mise à jour de la colonne 'TYPE'
    df_maxwire['TYPE'] = df_maxwire.apply(
        lambda row: 'Twist' if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '' and (pd.isna(row['TYPE']) or str(row['TYPE']).strip() == '') else row['TYPE'],
        axis=1
    )

    # Étape 5: Vérification des 'Twist'
    num_twist = df_maxwire[df_maxwire['TYPE'] == 'Twist'].shape[0]
    print(f"Count of 'Twist': {num_twist}")

    # Étape 6: Identification des lignes 'Twist' et non intégrées
    twist_df = df_maxwire[df_maxwire['TYPE'] == 'Twist'].copy()
    non_integrated_df2 = df_mmsta2_clean[~df_mmsta2_clean['SN2'].isin(twist_df['SN ADD'])]

    # Étape 7: Préparation des caractéristiques (conversion en chaînes de caractères)
    characteristics = ['From Seal by Terminal', 'From Terminal', 'To Terminal', 'To Seal by Terminal', 'Wire Part Number', 'Final Wire Length']

    for char in characteristics:
        df_maxwire.loc[:,char] = df_maxwire[char].astype(str).fillna('')

    non_integrated_df2.loc[:,'DS Général'] = non_integrated_df2['DS Général'].astype(str).fillna('')

    # Vérification de l'inclusion des caractéristiques
    def check_inclusion(row, ds_general):
        items = [row[char].strip() for char in characteristics if row[char].strip() not in ['', 'nan']]
        return all(item in ds_general for item in items)

    # Étape 8: Intégration de 'SN ADD' pour les lignes non intégrées
    def integrate_sn_add(row, non_integrated_df2):
        if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '':
            return row['SN ADD']  # Conserver la valeur existante si non vide
        matched_rows = non_integrated_df2[non_integrated_df2['DS Général'].apply(lambda x: check_inclusion(row, x))]
        if not matched_rows.empty:
            return matched_rows['SN2'].iloc[0]
        return ''  # Retourner une chaîne vide si aucune correspondance trouvée

    # Mettre à jour les valeurs de 'SN ADD' uniquement si elles sont vides
    df_maxwire['SN ADD'] = df_maxwire.apply(lambda row: integrate_sn_add(row, non_integrated_df2), axis=1)

    # Étape 9: Mettre à jour 'SN P2' là où 'TYPE' est 'Twist' et 'SN ADD' est égale à 'SN2' de df_mmsta2
    for i, row in df_maxwire[df_maxwire['TYPE'] == 'Twist'].iterrows():
        sn_add = row['SN ADD']
        sn_1_value = df_mmsta2.loc[df_mmsta2['SN2'] == sn_add, 'SN1']
        if not sn_1_value.empty:
            df_maxwire.at[i, 'SN P2'] = sn_1_value.values[0]



    # Étape 1: Nettoyage des données dans df_mmsta3 pour ignorer les SN2 vides ou contenant "(blanks)"
    df_mmsta3_clean = df_mmsta3.dropna(subset=['SN2']).copy()  # Supprime les lignes où SN2 est NaN
    df_mmsta3_clean = df_mmsta3_clean[df_mmsta3_clean['SN2'].str.strip() != '(blanks)']  # Supprime les lignes où SN2 est "(blanks)"

    # Étape 2: Remplissage de la colonne 'SN ADD' uniquement si elle est vide
    def map_sn_add(salma, df):
        matches = df[df['DSN2'] == salma]['SN2']
        return matches.iloc[0] if not matches.empty else ''

    # Conserver les valeurs existantes dans 'SN ADD' si elles ne sont pas vides
    def update_sn_add_if_empty(row, df_mmsta3_clean):
        if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '':
            return row['SN ADD']  # Conserver la valeur existante si non vide
        # Sinon, essayer de remplir 'SN ADD' avec la correspondance trouvée
        salma = row['salma']
        if salma in df_mmsta3_clean['DSN2'].values:
            return map_sn_add(salma, df_mmsta3_clean)
        return row.get('SN ADD', '')

    # Étape 3: Remplir la colonne 'SN ADD'
    df_maxwire['SN ADD'] = df_maxwire.apply(lambda row: update_sn_add_if_empty(row, df_mmsta3_clean), axis=1)

    # Étape 4: Mise à jour de la colonne 'TYPE'
    df_maxwire['TYPE'] = df_maxwire.apply(
        lambda row: 'SQUIB' if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '' and (pd.isna(row['TYPE']) or str(row['TYPE']).strip() == '') else row['TYPE'],
        axis=1
    )

    # Étape 5: Vérification des 'SQUIB'
    num_squib = df_maxwire[df_maxwire['TYPE'] == 'SQUIB'].shape[0]
    print(f"Count of 'SQUIB': {num_squib}")

    # Étape 6: Identification des lignes 'SQUIB' et non intégrées
    squib_df = df_maxwire[df_maxwire['TYPE'] == 'SQUIB'].copy()
    non_integrated_df3 = df_mmsta3_clean[~df_mmsta3_clean['SN2'].isin(squib_df['SN ADD'])]

    # Étape 7: Préparation des caractéristiques (conversion en chaînes de caractères)
    characteristics = ['From Seal by Terminal', 'From Terminal', 'To Terminal', 'To Seal by Terminal', 'Wire Part Number', 'Final Wire Length']

    for char in characteristics:
        df_maxwire.loc[:,char] = df_maxwire[char].astype(str).fillna('')

    non_integrated_df3.loc[:,'DS Général'] = non_integrated_df3['DS Général'].astype(str).fillna('')

    # Vérification de l'inclusion des caractéristiques
    def check_inclusion(row, ds_general):
        items = [row[char].strip() for char in characteristics if row[char].strip() not in ['', 'nan']]
        return all(item in ds_general for item in items)

    # Étape 8: Intégration de 'SN ADD' pour les lignes non intégrées
    def integrate_sn_add(row, non_integrated_df3):
        if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '':
            return row['SN ADD']  # Conserver la valeur existante si non vide
        matched_rows = non_integrated_df3[non_integrated_df3['DS Général'].apply(lambda x: check_inclusion(row, x))]
        if not matched_rows.empty:
            return matched_rows['SN2'].iloc[0]
        return ''  # Retourner une chaîne vide si aucune correspondance trouvée

    # Mettre à jour les valeurs de 'SN ADD' uniquement si elles sont vides
    df_maxwire['SN ADD'] = df_maxwire.apply(lambda row: integrate_sn_add(row, non_integrated_df3), axis=1)

    # Étape 9: Mettre à jour 'SN P2' là où 'TYPE' est 'SQUIB' et 'SN ADD' est égale à 'SN2' de df_mmsta3
    for i, row in df_maxwire[df_maxwire['TYPE'] == 'SQUIB'].iterrows():
        sn_add = row['SN ADD']
        sn_1_value = df_mmsta3.loc[df_mmsta3['SN2'] == sn_add, 'SN1']
        if not sn_1_value.empty:
            df_maxwire.at[i, 'SN P2'] = sn_1_value.values[0]


    # Étape 1: Nettoyage des données dans df_mmsta4 pour ignorer les SN2 vides ou contenant "(blanks)"
    df_mmsta4_clean = df_mmsta4.dropna(subset=['SN2']).copy()  # Supprime les lignes où SN2 est NaN
    df_mmsta4_clean = df_mmsta4_clean[df_mmsta4_clean['SN2'].str.strip() != '(blanks)']  # Supprime les lignes où SN2 est "(blanks)"

    # Étape 2: Remplissage de la colonne 'SN ADD' uniquement si elle est vide
    def map_sn_add(salma, df):
        matches = df[df['DSN2'] == salma]['SN2']
        return matches.iloc[0] if not matches.empty else ''

    # Conserver les valeurs existantes dans 'SN ADD' si elles ne sont pas vides
    def update_sn_add_if_empty(row, df_mmsta4_clean):
        if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '':
            return row['SN ADD']  # Conserver la valeur existante si non vide
        # Sinon, essayer de remplir 'SN ADD' avec la correspondance trouvée
        salma = row['salma']
        if salma in df_mmsta4_clean['DSN2'].values:
            return map_sn_add(salma, df_mmsta4_clean)
        return row.get('SN ADD', '')

    # Étape 3: Remplir la colonne 'SN ADD'
    df_maxwire['SN ADD'] = df_maxwire.apply(lambda row: update_sn_add_if_empty(row, df_mmsta4_clean), axis=1)

    # Étape 4: Mise à jour de la colonne 'TYPE'
    df_maxwire['TYPE'] = df_maxwire.apply(
        lambda row: 'GW' if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '' and (pd.isna(row['TYPE']) or str(row['TYPE']).strip() == '') else row['TYPE'],
        axis=1
    )

    # Étape 5: Vérification des 'GW'
    num_gw = df_maxwire[df_maxwire['TYPE'] == 'GW'].shape[0]
    print(f"Count of 'GW': {num_gw}")

    # Étape 6: Identification des lignes 'GW' et non intégrées
    gw_df = df_maxwire[df_maxwire['TYPE'] == 'GW'].copy()
    non_integrated_df4 = df_mmsta4_clean[~df_mmsta4_clean['SN2'].isin(gw_df['SN ADD'])]

    # Étape 7: Préparation des caractéristiques (conversion en chaînes de caractères)
    characteristics = ['From Seal by Terminal', 'From Terminal', 'To Terminal', 'To Seal by Terminal', 'Wire Part Number', 'Final Wire Length']

    for char in characteristics:
        df_maxwire.loc[:,char] = df_maxwire[char].astype(str).fillna('')

    non_integrated_df4.loc[:,'DS Général'] = non_integrated_df4['DS Général'].astype(str).fillna('')

    # Vérification de l'inclusion des caractéristiques
    def check_inclusion(row, ds_general):
        items = [row[char].strip() for char in characteristics if row[char].strip() not in ['', 'nan']]
        return all(item in ds_general for item in items)

    # Étape 8: Intégration de 'SN ADD' pour les lignes non intégrées
    def integrate_sn_add(row, non_integrated_df4):
        if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '':
            return row['SN ADD']  # Conserver la valeur existante si non vide
        matched_rows = non_integrated_df4[non_integrated_df4['DS Général'].apply(lambda x: check_inclusion(row, x))]
        if not matched_rows.empty:
            return matched_rows['SN2'].iloc[0]
        return ''  # Retourner une chaîne vide si aucune correspondance trouvée

    # Mettre à jour les valeurs de 'SN ADD' uniquement si elles sont vides
    df_maxwire['SN ADD'] = df_maxwire.apply(lambda row: integrate_sn_add(row, non_integrated_df4), axis=1)

    # Étape 9: Mettre à jour 'SN P2' là où 'TYPE' est 'GW' et 'SN ADD' est égale à 'SN2' de df_mmsta4
    for i, row in df_maxwire[df_maxwire['TYPE'] == 'GW'].iterrows():
        sn_add = row['SN ADD']
        sn_1_value = df_mmsta4.loc[df_mmsta4['SN2'] == sn_add, 'SN1']
        if not sn_1_value.empty:
            df_maxwire.at[i, 'SN P2'] = sn_1_value.values[0]




   

    # Étape 1: Nettoyage des données dans df_mmsta6 pour ignorer les SN3 vides ou contenant "(blanks)"
    df_mmsta6_clean = df_mmsta6.dropna(subset=['SN3']).copy()  # Supprime les lignes où SN3 est NaN
    df_mmsta6_clean = df_mmsta6_clean[df_mmsta6_clean['SN3'].str.strip() != '(blanks)']  # Supprime les lignes où SN3 est "(blanks)"

    # Étape 2: Fonction pour mapper 'SN ADD', 'SN P2' et 'Super group' en utilisant 'salma' et 'DSN3'
    def map_sn_add_sn_p2_and_supergroup(salma, df):
        matched_row = df[df['DSN3'] == salma]
        if not matched_row.empty:
            sn_add = matched_row['SN3'].tolist()
            sn_p2 = matched_row['SN2'].tolist()
            supergroup = matched_row['SN1'].tolist()
            return sn_add, sn_p2, supergroup
        return None, None, None

    # Fonction pour mettre à jour 'SN ADD', 'SN P2' et 'Super group'
    def update_sn_add_sn_p2_and_supergroup(row, df_mmsta6_clean):
        salma = row['salma']
        new_sn_add, new_sn_p2, new_supergroup = map_sn_add_sn_p2_and_supergroup(salma, df_mmsta6_clean)

        # Mise à jour de SN ADD
        if isinstance(row['SN ADD'], list) and any(str(val).strip() for val in row['SN ADD']):
            sn_add_value = row['SN ADD']
        else:
            sn_add_value = new_sn_add if new_sn_add else row['SN ADD']

        # Mise à jour de SN P2
        if isinstance(row['SN P2'], list) and any(str(val).strip() for val in row['SN P2']):
            sn_p2_value = row['SN P2']
        else:
            sn_p2_value = new_sn_p2 if new_sn_p2 else row['SN P2']

        # Mise à jour de Super group
        if isinstance(row['Super group'], list) and any(str(val).strip() for val in row['Super group']):
            supergroup_value = row['Super group']
        else:
            supergroup_value = new_supergroup if new_supergroup else row['Super group']

        return sn_add_value, sn_p2_value, supergroup_value

    # Étape 3: Remplir les colonnes 'SN ADD', 'SN P2', et 'Super group'
    df_maxwire[['SN ADD', 'SN P2', 'Super group']] = df_maxwire.apply(lambda row: update_sn_add_sn_p2_and_supergroup(row, df_mmsta6_clean), axis=1, result_type='expand')

    # Étape 4: Aplatir les listes des SN ADD, SN P2, et Super group
    df_maxwire['SN ADD'] = df_maxwire['SN ADD'].apply(lambda x: x if isinstance(x, list) else [x] if pd.notna(x) else [])
    df_maxwire['SN P2'] = df_maxwire['SN P2'].apply(lambda x: x if isinstance(x, list) else [x] if pd.notna(x) else [])
    df_maxwire['Super group'] = df_maxwire['Super group'].apply(lambda x: x if isinstance(x, list) else [x] if pd.notna(x) else [])

    df_maxwire = df_maxwire.explode(['SN ADD', 'SN P2', 'Super group']).reset_index(drop=True)

    # Étape 5: Remplissage de la colonne 'TYPE'
    df_maxwire['TYPE'] = df_maxwire.apply(
        lambda row: 'SG' if pd.notna(row['SN ADD']) and str(row['SN ADD']).strip() != '' and (pd.isna(row['TYPE']) or str(row['TYPE']).strip() == '') else row['TYPE'],
        axis=1
    )

    # Étape 6: Mise à jour des colonnes de df_maxwire en fonction des colonnes de df_mmsta6
    columns_to_check = df_mmsta6.columns[df_mmsta6.columns.get_loc('DSN3'):df_mmsta6.columns.get_loc('Total') + 1]
    columns_to_update = df_maxwire.columns[df_maxwire.columns.get_loc('To Eyelet Protection'):df_maxwire.columns.get_loc('salma') + 1]

    for index, row in df_maxwire.iterrows():
        salma = row['salma']
        matched_row = df_mmsta6_clean[df_mmsta6_clean['DSN3'] == salma]

        if not matched_row.empty:
            for col_to_check, col_to_update in zip(columns_to_check, columns_to_update):
                if matched_row.iloc[0][col_to_check] == 1:
                    df_maxwire.loc[index, col_to_update] = 'X'
                else:
                    df_maxwire.loc[index, col_to_update] = np.nan  # Supprimer le "X" s'il ne doit pas être présent

    # Suppression des lignes dupliquées où TYPE='SG' et la combinaison SN ADD, SN P2, et Super group est identique
    df_maxwire = df_maxwire[~((df_maxwire['TYPE'] == 'SG') & df_maxwire.duplicated(subset=['SN ADD', 'SN P2', 'Super group'], keep='first'))]

    # Vérification du résultat final
    print(df_maxwire.head())

    # Suppression des lignes dupliquées dans df_maxwire
    #df_maxwire_cleaned = df_maxwire.drop_duplicates()


    # Étape 12: Sauvegarde et affichage
    output_path = 'liste_circuit_integre.xlsx'
    
    # Création d'un ExcelWriter pour sauvegarder plusieurs DataFrames dans différentes feuilles
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sauvegarde de df_maxwire dans une feuille
        df_maxwire.to_excel(writer, sheet_name='Integrated Data', index=False)

        # Comptage du nombre de lignes pour chaque type dans df_maxwire
        types = ['FIL SIMPLE', 'Double', 'Twist', 'SQUIB', 'GW', 'Joint', 'SG']
        type_counts_maxwire = {t: len(df_maxwire[df_maxwire['TYPE'] == t]) for t in types}

        # Comptage du nombre de lignes dans les colonnes 'DS Général' de chacun des df_mmsta
        type_counts_mmsta = {
            'FIL SIMPLE': df_mmsta['DS Général'].dropna().count(),
            'Double': df_mmsta1['DS Général'].dropna().count(),
            'Twist': df_mmsta2['DS Général'].dropna().count(),
            'SQUIB': df_mmsta3['DS Général'].dropna().count(),
            'GW': df_mmsta4['DS Général'].dropna().count(),
            'Joint': df_mmsta5['DS Général'].dropna().count(),
            'SG': df_mmsta6['DS Général'].dropna().count()
        }

        # Création d'un tableau comparatif
        comparison_df = pd.DataFrame({
            'Type': types,
            'Nombre de lignes dans df_mmsta': [type_counts_mmsta[t] for t in types],
            'Nombre de lignes dans df_maxwire': [type_counts_maxwire[t] for t in types]
        })

        # Sauvegarde du tableau comparatif dans une feuille
        comparison_df.to_excel(writer, sheet_name='Comparison', index=False)

    # Affichage du tableau comparatif
    print(comparison_df)

    return output_path

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python integration.py <separated_file_path> <circuit_file_path>")
        sys.exit(1)
    
    separated_file = sys.argv[1]
    circuit_file = sys.argv[2]
    output_file = integrate_files(separated_file, circuit_file)
    
    if output_file:
        print(f"Excel file generated: {output_file}")
    else:
        print("No output file generated.")

