import pandas as pd 
import json
import numpy as np
from datetime import datetime
from datetime import time

def convert_birchstreet_excel_to_json(file_path1, file_path2, output_file_path1, output_file_path2, header_row_index1, header_row_index2):
    """
    Converts a pair of Excel files to JSON files, considering the header row index for each file.

    :param file_path1: str, path to the first Excel file.
    :param file_path2: str, path to the second Excel file.
    :param output_file_path1: str, path where the first json file will be saved.
    :param output_file_path2: str, path where the second json file will be saved.
    :param header_row_index1: int, the index of the row containing the header in the first file.
    :param header_row_index2: int, the index of the row containing the header in the second file.
    """
     # Charger le premier fichier Excel en spécifiant l'index de la ligne d'en-tête
    df1 = pd.read_excel(file_path1, header=5)
    # Convertir le DataFrame en JSON avec l'orientation "records"
    with open(output_file_path1, 'w') as f:
        json.dump(df1.to_dict(orient='records'), f, indent=4)

    # Charger le deuxième fichier Excel de la même manière
    df2 = pd.read_excel(file_path2, header=5)
    # Convertir et sauvegarder le deuxième DataFrame en JSON
    with open(output_file_path2, 'w') as f:
        json.dump(df2.to_dict(orient='records'), f, indent=4)

    return f"Files converted and saved as '{output_file_path1}' and '{output_file_path2}'"

def convert_checkscm_excel_to_json(file_path1, file_path2, file_path3,output_file_path1, output_file_path2,output_file_path3):
    """
    Converts three Excel files to JSON files

    :param file_path1: str, path to the first Excel file.
    :param file_path2: str, path to the second Excel file.
    :param file_path3: str, path to the third Excel file.
    :param output_file_path1: str, path where the first json file will be saved.
    :param output_file_path2: str, path where the second json file will be saved.
    :param output_file_path3: str, path where the third json file will be saved.
    """
    
    df1 = pd.read_excel(file_path1)
    with open(output_file_path1, 'w') as f:
        json.dump(df1.to_dict(orient='records'), f, indent=4)

    df2 = pd.read_excel(file_path2)
    with open(output_file_path2, 'w') as f:
        json.dump(df2.to_dict(orient='records'), f, indent=4)

    df3 = pd.read_excel(file_path3)
    with open(output_file_path3, 'w') as f:
        json.dump(df3.to_dict(orient='records'), f, indent=4)

    return f"Files converted and saved as '{output_file_path1}' ,'{output_file_path2}' and '{output_file_path3}"

def time_to_float(t):
    # t est un objet datetime.time
    return (t.hour * 3600 + t.minute * 60 + t.second) / 86400.0  # 86400 secondes dans une journée

def convert_iscala_excel_to_json(file_paths, output_file_paths):
    """
    Converts Excel files to JSON files.

    :param file_paths: list of str, paths to the Excel files.
    :param output_file_paths: list of str, paths where the JSON files will be saved.
    """
    for file_path, output_file_path in zip(file_paths, output_file_paths):
        list_products = []
        current_product = {}
        list_ingredients = []
        
        # Lecture du fichier Excel, suppression des lignes NaN et des lignes spéciales
        df = pd.read_excel(file_path, sheet_name='Sheet1', skiprows=6, header=None)
        df = df.dropna(axis=0, how='all')
        df = df[~df[0].isin(['- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'])]
        
        # Traitement des lignes du DataFrame
        for index, row in df.iterrows():
            row = row.values  # Convertir la série Pandas en tableau numpy pour un accès plus facile
            # Gestion des templates
            if pd.notnull(row[0]) and str(row[0]).startswith('Template : '):
                if current_product.get("name"):
                    current_product['ingredients'] = list_ingredients
                    list_products.append(current_product)
                    list_ingredients = []
                current_product = {"name": str(row[0])[11:]}
                continue

            # Gestion des totaux
            if pd.notnull(row[0]) and str(row[0]) == 'Total for Stock Item':
                current_product['total'] = {
                    'Sell price': row[11],
                    'Sell': row[17],
                    'Cost': str(row[21]),
                    'Cost%': row[24]
                }
                current_product['ingredients'] = list_ingredients
                list_products.append(current_product)
                current_product = {}
                list_ingredients = []
                continue

            # Gestion des ingrédients
            list_ingredients.append({
                'code': row[0] if pd.notnull(row[0]) else '000',
                'name': row[8] if pd.notnull(row[8]) else row[0],
                'L': row[14] if pd.notnull(row[14]) else '000',
                'S': row[20] if pd.notnull(row[20]) else '000',
                'Qty': row[24] if pd.notnull(row[24]) else '000',
                'Unit': row[32] if pd.notnull(row[32]) else '000',
                'Sell_price': row[36] if pd.notnull(row[36]) else '000',
                'Cost%': row[42] if pd.notnull(row[42]) else '000'
            })

        # Assurer que le dernier produit est ajouté s'il n'est pas vide
        if current_product.get("name"):
            current_product['ingredients'] = list_ingredients
            list_products.append(current_product)

        # Création du fichier JSON de sortie
        with open(output_file_path, "w") as outfile:
            json.dump(list_products, outfile, indent=4)

    return "Files converted and saved."


def convert_manual_to_json(file_path, json_output_path):
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names[1:]  # Ignorer la première feuille
    all_extracted_data = []

    for sheet_name in sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        # Extraction des informations spécifiques de la feuille
        name_pizza = df.iloc[1, 0]
        cout_portion_pizza = df.iloc[3, 6]
        sell_price_without_tax_pizza = df.iloc[4, 6]
        sell_price_with_tax_pizza = df.iloc[5, 6]
        cost_percent_pizza = df.iloc[6, 6]
        
        extracted_values_pizza = {
            "Name": name_pizza,
            "Cout Portion": cout_portion_pizza,
            "Sell Price without Tax": sell_price_without_tax_pizza if pd.notnull(sell_price_without_tax_pizza) else '000',
            "Sell Price with Tax": sell_price_with_tax_pizza if pd.notnull(sell_price_with_tax_pizza) else '000',
            "Cost%": cost_percent_pizza if pd.notnull(cost_percent_pizza) else '000',
            "Ingredient": []
        }
        
        start_row_index = 9
        extracted_info = []
        
        for row in df.iloc[start_row_index:].itertuples(index=False):
            if any("direction" in str(cell).lower() for cell in row):
                break
            else:
                extracted_info.append(row)
        
        formatted_data = []
        for row in extracted_info:
            if row[0] is None or str(row[0]).strip() == '':
                continue
            formatted_row = {
                "code": row[2] if pd.notnull(row[2]) else '000',
                "Nom": row[0] if pd.notnull(row[0]) else '000',
                "Quantity": row[3] if pd.notnull(row[3]) else '000',
                "Quantity2": row[4] if pd.notnull(row[4]) else '000',
                "Unit": row[5] if pd.notnull(row[5]) else '000',
                "Cost per unit": row[6] if pd.notnull(row[6]) else '000',
                "Cost per ingredient": row[7] if pd.notnull(row[7]) else '000'
            }
            formatted_data.append(formatted_row)
        
        for item in formatted_data:
            extracted_values_pizza['Ingredient'].append(item)
        
        all_extracted_data.append(extracted_values_pizza)

    # Écrire les données extraites dans un fichier JSON
    with open(json_output_path, 'w') as json_file:
        json.dump(all_extracted_data, json_file, indent=4, default=str)

def process_manual_files(excel_files, output_files):
    for file_path, output_path in zip(excel_files, output_files):
        convert_manual_to_json(file_path, output_path)
        print(f"Converted {file_path} to {output_path}")

def process_birchstreet_files(beverage_recipe_path, ingredient_masterlist_path, category_ingredient_path, agbcode_file_path, sales_analysis_path, output_path):
    # Lecture des fichiers Excel
    b_r_birch = pd.read_excel(beverage_recipe_path, header=5)
    b_i_birch = pd.read_excel(ingredient_masterlist_path)
    c_i_birch = pd.read_excel(category_ingredient_path)
    agbsrc = pd.read_excel(agbcode_file_path)
    agbsrc.head()
    sales_a = pd.read_excel(sales_analysis_path)

    # Filtrage des données
    b_r_birch = b_r_birch[b_r_birch['Recipe type'] == 'Recipe']
    agbsrc = agbsrc[['Code\nAGB', 'WiseFins FR Category', 'WiseFins EN Category', 'WiseFins FR Subcategory', 'WiseFins EN Subcategory']]
    agbsrc.drop_duplicates(inplace=True)
    agbsrc['Code AGB'] = agbsrc['Code\nAGB'].astype(str)

    # Ajout de colonnes supplémentaires
    additional_columns = {
        'resto': 'birchstreet',
        'recipeGroup': 'Beverage',
        'subRecipe': np.nan,
        'statut': 'ACTIVE',
        'photo': 'no',
        'portion': 1
    }
    b_r_birch = b_r_birch.assign(**additional_columns)

    # Renommage de colonnes
    column_rename_mapping = {
        'Ingredient/Subrecipe name': 'recipeCompose',
        'Dish/Recipe name': 'recipeName',
        'UOM': 'unit'
    }
    b_r_birch.rename(columns=column_rename_mapping, inplace=True)

    # Ajout de la colonne 'quantityAfter'
    b_r_birch['quantityAfter'] = b_r_birch['Quantity'].copy()

    # Traitement des codes SKU
    b_r_birch['SKU'] = b_r_birch['recipeCompose'].apply(lambda x: x[:8])
    b_r_birch['SKU'] = b_r_birch['SKU'].astype(str)
    b_i_birch['Supplier SKU'] = b_i_birch['Supplier SKU'].astype(str)

    # Calcul du prix unitaire moyen
    b_i_birch = b_i_birch.groupby('Supplier SKU', as_index=False)['Unit price'].mean()
    b_i_birch['Supplier SKU'] = b_i_birch['Supplier SKU'].astype(str)

    # Fusion avec les données des ingrédients Birch Street
    b_r_birch = b_r_birch.merge(b_i_birch[['Supplier SKU', 'Unit price']], how='left', left_on='SKU', right_on='Supplier SKU')
    b_r_birch.drop_duplicates(inplace=True)
    b_r_birch.rename(columns={'Unit price': 'cost'}, inplace=True)

    # Traitement des codes AGB
    c_i_birch = c_i_birch[['Part #', 'AGBCode']]
    c_i_birch.drop_duplicates(inplace=True)
    b_r_birch = b_r_birch.merge(c_i_birch[['AGBCode', 'Part #']], how='left', left_on='SKU', right_on='Part #')
    b_r_birch['AGBCode'].fillna(0, inplace=True)
    b_r_birch['AGBCode'] = b_r_birch['AGBCode'].astype(str)

    # Fusion avec les données AGB pour obtenir les catégories
    b_r_birch = b_r_birch.merge(agbsrc[['Code AGB', 'WiseFins EN Category', 'WiseFins EN Subcategory']], how='left', left_on='AGBCode', right_on='Code AGB')
    b_r_birch.rename(columns={'WiseFins EN Category': 'category', 'WiseFins EN Subcategory': 'subcategory'}, inplace=True)

    # Traitement des données de vente
    sales_a['code_grpparent'] = sales_a['Code'].apply(lambda x: x[-8:])
    b_r_birch['code_grp'] = b_r_birch['recipeName'].apply(lambda x: x[-8:])
    sales_a['code_grpparent'] = sales_a['code_grpparent'].astype(str)
    b_r_birch['code_grp'] = b_r_birch['code_grp'].astype(str)
    b_r_birch = b_r_birch.merge(sales_a[['Sales Price', 'recipeGroupParent', 'code_grpparent']], how='left', left_on='code_grp', right_on='code_grpparent')
    b_r_birch.rename(columns={'Sales Price': 'selling'}, inplace=True)
    b_r_birch['selling'] = b_r_birch['selling'].fillna(0.0001)

    # Finalisation et exportation des données
    b_r_birch.rename(columns={'SKU': 'InventoryCode'}, inplace=True)
    b_r_birch['recipeName'] = b_r_birch['recipeName'].apply(lambda x: x[10:-8].strip())
    b_r_birch['AGBCode'] = b_r_birch['AGBCode'].astype(int)
    column_order = ['resto', 'subRecipe', 'recipeGroup', 'recipeGroupParent', 'recipeName', 'recipeCompose', 'AGBCode', 'Quantity', 'cost', 'InventoryCode', 'unit', 'quantityAfter', 'portion', 'statut', 'photo', 'selling', 'category', 'subcategory']
    b_r_birch = b_r_birch[column_order]
    b_r_birch.to_excel(output_path, index=False)


def transform(file, output_file_path):
    # Charger le fichier Excel dans un DataFrame
    sales_a = pd.read_excel(file)
    
    # Obtenir le nombre total de lignes dans le DataFrame
    size = sales_a.shape[0]
    
    # Parcourir chaque ligne pour trouver la fin du rapport
    for i in range(size):
        if str(sales_a.iloc[i, 0]).startswith("** END OF REPORT **"):
            size = i  # Mettre à jour la taille pour exclure tout ce qui suit cette ligne
            break  # Sortir de la boucle une fois la fin du rapport trouvée

    i = 0
    data = pd.DataFrame()  # Initialiser un DataFrame vide pour stocker les données nettoyées

    # Parcourir les lignes du DataFrame pour traiter les données de chaque catégorie
    while i < size:
        # Vérifier si la ligne actuelle marque le début d'une nouvelle catégorie
        if str(sales_a.iloc[i, 0]).startswith("Category: "):
            header = sales_a.iloc[i + 1, :]  # La ligne suivante contient les en-têtes de colonne
            
            # Parcourir les lignes suivantes pour extraire les données de cette catégorie
            for j in range(i + 2, size):
                # Vérifier si on a atteint le sous-total de la catégorie
                if str(sales_a.iloc[j, 0]).startswith("Sub Total:"):
                    cat = sales_a.iloc[i]  # Extraire le nom de la catégorie
                    cat = cat[0].split("Category: ")  # Nettoyer le nom de la catégorie
                    parent_grp = cat[1]  # Stocker le nom de la catégorie parente
                    
                    # Extraire les données de vente de la catégorie et ajouter une colonne pour le groupe parent
                    sales_data = sales_a.iloc[i + 2:j, :].copy()
                    sales_data.loc[:, 'recipeGroupParent'] = parent_grp
                    
                    # Concaténer les données extraites avec le DataFrame principal
                    data = pd.concat([data, sales_data], axis=0)
                    
                    i = j + 3  # Mettre à jour l'index pour passer au traitement de la prochaine catégorie
                    break  # Sortir de la boucle interne pour commencer à traiter la prochaine catégorie
        else:
            i += 1  # Passer à la ligne suivante si la ligne actuelle ne marque pas le début d'une nouvelle catégorie

    # Remplacer les chaînes vides par NaN et supprimer les colonnes entièrement NaN
    data.replace("", float("NaN"), inplace=True)
    data.dropna(how='all', axis=1, inplace=True)

    # Nettoyer l'en-tête et créer une liste pour les noms de colonnes
    header = header.dropna()
    header_list = header.tolist()
    header_list.append('recipeGroupParent')

    # Appliquer les noms de colonnes au DataFrame et ajouter la colonne du groupe parent
    data.columns = header_list

    # Exporter les données nettoyées vers un nouveau fichier Excel, en utilisant le chemin fourni
    data.to_excel(output_file_path, index=False)

def convert_price_to_float(price_str):
    # Supprime les virgules et convertit en float
    return float(price_str.replace(',', ''))


def create_excel_manual_from_json_file(json_file_path, output_excel_path):
    # Load JSON data from the file
    with open(json_file_path, 'r') as file:
        json_data = json.load(file)

    # Initialize an empty list to store the DataFrame rows
    rows = []

    # Loop through each entry in the JSON data and create a row for each ingredient
    for item in json_data:
        # Determine if the item is food or beverage based on the first letter of the name
        first_letter = item['Name'].strip()[0].upper()
        recipe_group = 'food' if first_letter == 'F' else 'beverage' if first_letter == 'B' else 'other'

        for ingredient in item['Ingredient']:
            
            cost = item['Sell Price with Tax'] if item['Cost%'] == 0 else item['Sell Price with Tax'] * item['Cost%']

            # Create a dictionary for the row
            row = {
                'resto': 'Manual',
                'subRecipe': '',
                'recipeGroup': recipe_group,
                'recipeGroupParent': '',
                'recipeName': item['Name'].strip(),
                'recipeCompose': ingredient['Nom'],
                'AGBCode': '',
                'Quantity': ingredient['Quantity'],
                'cost': cost,
                'InventoryCode': ingredient['code'],
                'unit': ingredient['Unit'],
                'quantityAfter': '',
                'portion': 1,
                'statut': '',
                'photo': '',
                'selling': item['Sell Price with Tax'],
                'category': '',
                'subcategory': ''
            }

            # Append the row dictionary to the rows list
            rows.append(row)

    # Convert the list of rows to a DataFrame
    df = pd.DataFrame(rows)

    # Save the DataFrame to an Excel file
    df.to_excel(output_excel_path, index=False)

def create_excel_from_json_file(json_file_path, output_excel_path):
    # Load JSON data from the file
    with open(json_file_path, 'r') as file:
        json_data = json.load(file)

    # Initialize an empty list to store the DataFrame rows
    rows = []

    # Loop through each entry in the JSON data and create a row for each ingredient
    for item in json_data:
        # Determine if the item is food or beverage based on the first letter of the name
        first_letter = item['name'].strip()[0].upper()
        recipe_group = 'food' if first_letter == 'F' else 'beverage' if first_letter == 'B' else 'other'

        for ingredient in item['ingredients']:
            # Calculate the cost by dividing the sell price by the cost percentage
            try:
                sell_price = convert_price_to_float(ingredient['Sell_price'].replace(',', '').replace('R', ''))
            except ValueError:
                sell_price = convert_price_to_float(ingredient['Cost%'].strip('%')) if ingredient['Cost%'].strip('%') else 0

            # Calculate cost if possible
            cost_percentage = convert_price_to_float(ingredient['Cost%'].strip('%')) / 100 if ingredient['Cost%'].strip('%') else 0
            cost = sell_price if cost_percentage == 0 else sell_price * cost_percentage

            # Create a dictionary for the row
            row = {
                'resto': 'iscala',
                'subRecipe': '',
                'recipeGroup': recipe_group,
                'recipeGroupParent': '',
                'recipeName': item['name'],
                'recipeCompose': ingredient['name'],
                'AGBCode': '',
                'Quantity': ingredient['Qty'],
                'cost': cost,
                'InventoryCode': ingredient['code'],
                'unit': ingredient['Unit'],
                'quantityAfter': '',
                'portion': 1,
                'statut': '',
                'photo': '',
                'selling': ingredient['Sell_price'],
                'category': '',
                'subcategory': ''
            }

            # Append the row dictionary to the rows list
            rows.append(row)

    # Convert the list of rows to a DataFrame
    df = pd.DataFrame(rows)

    # Save the DataFrame to an Excel file
    df.to_excel(output_excel_path, index=False)

def create_excel_checkscm_ingre_from_json_file(json_file_path, output_excel_path):
    # Load JSON data from the file
    with open(json_file_path, 'r') as file:
        json_data = json.load(file)

    # Initialize an empty list to store the DataFrame rows
    rows = []

    # Loop through each entry in the JSON data and create a row for each ingredient
    for item in json_data:
        # Determine if the item is food or beverage based on the first letter of the name
        first_letter = item['Department'].strip()[0].upper()
        recipe_group = 'food' if first_letter == 'F' else 'beverage' if first_letter == 'B' else 'other'
        row = {
                'resto': 'Checkscm',
                'subRecipe': item['Type'],
                'recipeGroup': item['Group'],
                'recipeGroupParent': recipe_group,
                'recipeName': item['Product Description'].strip(),
                'recipeCompose': '',
                'AGBCode': '',
                'Quantity': '',
                'cost': item['Avg. Price'],
                'InventoryCode': item['Product'],
                'unit': item['Stock Size'],
                'quantityAfter': '',
                'portion': 1,
                'statut': '',
                'photo': '',
                'selling': '',
                'category': '',
                'subcategory': ''
        }
        rows.append(row)

    # Convert the list of rows to a DataFrame
    df = pd.DataFrame(rows)

    # Save the DataFrame to an Excel file
    df.to_excel(output_excel_path, index=False)

def create_excel_checkscm_recipe_from_json_file(json_file_path, output_excel_path):
    # Load JSON data from the file
    with open(json_file_path, 'r') as file:
        json_data = json.load(file)

    # Initialize an empty list to store the DataFrame rows
    rows = []

    # Loop through each entry in the JSON data and create a row for each ingredient
    for item in json_data:
        # Determine if the item is food or beverage based on the first letter of the name
        first_letter = item['Recipe Group'].strip()[0].upper()
        recipe_group = 'food' if first_letter == 'F' else 'beverage' if first_letter == 'B' else 'other'
        row = {
                'resto': 'Checkscm',
                'subRecipe': item['Class'],
                'recipeGroup': item['Recipe Type'],
                'recipeGroupParent': recipe_group,
                'recipeName': item['Recipe Description'].strip(),
                'recipeCompose': '',
                'AGBCode': '',
                'Quantity':'' ,
                'cost': item['Total Cost'],
                'InventoryCode': item['Recipe Number'],
                'unit': item['Serve Size'],
                'quantityAfter': '',
                'portion': item['No. of Serve'],
                'statut': item['Status'],
                'photo': '',
                'selling': item['Suggested Selling Price'],
                'category': '',
                'subcategory': ''
        }
        rows.append(row)

    # Convert the list of rows to a DataFrame
    df = pd.DataFrame(rows)

    # Save the DataFrame to an Excel file
    df.to_excel(output_excel_path, index=False)


def main():

    ######################################################################################################
    ##                             Récupération des fichiers dans les répertoires                       ##
    #######################################################################################################

    # paths for birchstreet
    beverage_birchstreet_path = 'data\\20231207_client_data\\export from Purchasing software\\BIRCHSTREET (Hotel 1)\\beverage_recipe_birchstreet.xlsx'
    ingredient_birchstreet_path = 'data\\20231207_client_data\\export from Purchasing software\BIRCHSTREET (Hotel 1)\\ingredient_masterlist_birchsreet.xlsx'
    # paths for checkscm
    beverage_checkscm_path = 'data\\20231207_client_data\\export from Purchasing software\\CHECKSCM (Hotel 3)\\beverage_ingredient_masterlist_checkscm.xlsx'
    ingredient_checkscm_path = 'data\\20231207_client_data\\export from Purchasing software\\CHECKSCM (Hotel 3)\\food_ingredient_masterlist_checkscm.xlsx'
    recipe_checkscm_path = 'data\\20231207_client_data\\export from Purchasing software\\CHECKSCM (Hotel 3)\\food_recipe_checkscm.xlsx'
    # Output paths for the JSON files
    # output paths for birchstreet
    output_json_birchstreet_beverage_path = 'data_converted\\birchstreet\\beverage_recipe_birchstreet.json'
    output_json_birchstreet_ingredient_path = 'data_converted\\birchstreet\ingredient_masterlist_birchstreet.json'
    # output paths for checkscm
    output_json_checkscm_beverage_path = 'data_converted\\checkscm\\beverage_ingredient_masterlist_checkscm.json'
    output_json_checkscm_ingredient_path = 'data_converted\\checkscm\\food_ingredient_masterlist_checkscm.json'
    output_json_checkscm_recipe_path = 'data_converted\\checkscm\\food_recipe_checkscm.json'
    
    # Identified header row indices for both files
    header_row_index_beverage = 5  
    header_row_index_ingredient = 0  

    # #path for manual data
    manual_excel_files = [
        "data\\20231207_client_data\\recipe manual input on excel\\Test\\chef_manual_recipe_conversion_hotel1.xlsx",
        "data\\20231207_client_data\\recipe manual input on excel\\Test\\chef_manual_recipe_pizza_hotel1.xlsx",
    ]
    output_manual_excel_files = [
                                 "data_converted\\manual\\chef_manual_recipe_conversion_hotel1.json", 
                                 "data_converted\\manual\\chef_manual_recipe_pizza_hotel1.json"
    ]
    
    # paths for iscala
    file_iscala_paths = ['data\\20231207_client_data\\export from Purchasing software\\ISCALA (Hotel 2)\\food_recipe_iscala.xlsx', 
              'data\\20231207_client_data\\export from Purchasing software\ISCALA (Hotel 2)\\beverage_recipe_iscala.xlsx',
                'data\\20231207_client_data\\export from Purchasing software\\ISCALA (Hotel 2)\\conversion_recipe_iscala.xlsx']


    #    output paths for iscala
    output_file__iscala_paths = ["data_converted\\iscala\\food_recipe_iscala.json", 
                             "data_converted\\iscala\\beverage_recipe_iscala.json", 
                             "data_converted\\iscala\\conversion_recipe_iscala.json"]
    

    

    json_file_iscala_paths = [
        #"data_converted\\iscala\\food_recipe_iscala.json", 
        "data_converted\\iscala\\beverage_recipe_iscala.json",
    ]
    excel_output_iscala_paths = [
        #'data_output\\iscala\\food_recipe_iscala.xlsx',
        'data_output\\iscala\\beverage_recipe_iscala.xlsx',
    ]

    input_manual_json_files = [
                                #  "data_converted\\manual\\chef_manual_recipe_conversion_hotel1.json", 
                                 "data_converted\\manual\\chef_manual_recipe_pizza_hotel1.json"
    ]

    excel_output_manual_files = [
                                #  "data_output\\manual\\chef_manual_recipe_conversion_hotel1.xlsx", 
                                 "data_output\\manual\\chef_manual_recipe_pizza_hotel1.xlsx"
    ]

    input_checkscm_ingre_json_files = [
                                 "data_converted\\checkscm\\beverage_ingredient_masterlist_checkscm.json", 
                                 "data_converted\\checkscm\\food_ingredient_masterlist_checkscm.json"
    ]

    excel_checkscm_ingre_files = [
                                 "data_output\\checkscm\\beverage_ingredient_masterlist_checkscm.xlsx", 
                                 "data_output\\checkscm\\food_ingredient_masterlist_checkscm.xlsx"
    ]

    input_checkscm_recipe_json_files = [
                                 "data_converted\\checkscm\\food_recipe_checkscm.json"
    ]

    excel_checkscm_recipe_files = [
                                 "data_output\\checkscm\\food_recipe_checkscm.xlsx"
    ]

    transform_file = "data\menu_sales_analysis_pos.xlsx"
    output_transform_file = "data_converted\\transform\\output_menu_sales_analysis_transformed.xlsx"


    ######################################################################################################
    ##                                          Appel de fonction                                       ##
    #######################################################################################################

    convert_birchstreet_excel_to_json(beverage_birchstreet_path, ingredient_birchstreet_path, 
                              output_json_birchstreet_beverage_path, output_json_birchstreet_ingredient_path, 
                              header_row_index_beverage, header_row_index_ingredient)

    convert_checkscm_excel_to_json(beverage_checkscm_path,ingredient_checkscm_path,recipe_checkscm_path,
                                  output_json_checkscm_beverage_path,output_json_checkscm_ingredient_path,output_json_checkscm_recipe_path)



    convert_iscala_excel_to_json(file_iscala_paths,output_file__iscala_paths)

    process_manual_files(manual_excel_files, output_manual_excel_files)

    transform(transform_file,output_transform_file)

    for json_path, excel_path in zip(json_file_iscala_paths, excel_output_iscala_paths):
        create_excel_from_json_file(json_path, excel_path)

    

    # # Calling the function to convert both Excel files to CSV
    

    for json_path, excel_path in zip(input_checkscm_ingre_json_files, excel_checkscm_ingre_files):
        create_excel_checkscm_ingre_from_json_file(json_path, excel_path)

    for json_path, excel_path in zip(input_checkscm_recipe_json_files, excel_checkscm_recipe_files):
        create_excel_checkscm_recipe_from_json_file(json_path, excel_path)
    

    for json_path, excel_path in zip(input_manual_json_files, excel_output_manual_files):
        create_excel_manual_from_json_file(json_path, excel_path)

    

    process_birchstreet_files(
        'data\\20231207_client_data\\export from Purchasing software\\BIRCHSTREET (Hotel 1)\\beverage_recipe_birchstreet.xlsx',
        'data\\20231207_client_data\\export from Purchasing software\BIRCHSTREET (Hotel 1)\\ingredient_masterlist_birchsreet.xlsx',
        'data\\20231207_client_data\\export from Purchasing software\BIRCHSTREET (Hotel 1)\\birchstreet_reference_files\\Copy of Category_ingredient_BirchStreet.xlsx',
        'data\\Agribalyse 2023 3.1 with WiseFins categorisation.xlsx',
        'data_converted\\transform\\output_menu_sales_analysis_transformed.xlsx',
        'data_output\\birchstreet\\output_birchstreet.xlsx'
    )

if __name__ == "__main__":
    main()