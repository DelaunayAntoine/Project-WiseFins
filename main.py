import pandas as pd 
import json
import numpy as np

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
    # Load the first Excel file, specifying the header row index
    df1 = pd.read_excel(file_path1, header=header_row_index1)
    # Save the first dataframe to a CSV file
    df1.to_json(output_file_path1, orient="index")


    df2 = pd.read_excel(file_path2, header=header_row_index2)
    df2.to_json(output_file_path2, orient="index")

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
    df1.to_json(output_file_path1, orient="index")

    df2 = pd.read_excel(file_path2)
    df2.to_json(output_file_path2, orient="index")

    df3 = pd.read_excel(file_path3)
    df3.to_json(output_file_path3, orient="index")

    return f"Files converted and saved as '{output_file_path1}' ,'{output_file_path2}' and '{output_file_path3}"

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
        for _, row in df.iterrows():
            if pd.notnull(row[1]) and str(row[1]).startswith('Template : '):
                if current_product.get("name"):
                    current_product['ingredients'] = list_ingredients
                    list_products.append(current_product)
                    list_ingredients = []
                current_product = {"name": row[1][11:]}
                continue

            if pd.notnull(row[1]) and str(row[1]) == 'Total for Stock Item':
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

            list_ingredients.append({
                'code': row[0] if pd.notnull(row[0]) else '000',
                'name': row[8] if pd.notnull(row[8]) else '000',
                'L': row[14] if pd.notnull(row[14]) else '000',
                'S': row[20] if pd.notnull(row[20]) else '000',
                'Qty': row[24] if pd.notnull(row[24]) else '000',
                'Unit': row[32] if pd.notnull(row[32]) else '000',
                'Sell_price': row[36] if pd.notnull(row[36]) else '000',
                'Cost%': row[42] if pd.notnull(row[42]) else '000'
            })

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

# Paths to the Excel files
# paths for birchstreet
beverage_birchstreet_path = 'data\\20231207_client_data\\export from Purchasing software\\BIRCHSTREET (Hotel 1)\\beverage_recipe_birchstreet.xlsx'
ingredient_birchstreet_path = 'data\\20231207_client_data\\export from Purchasing software\BIRCHSTREET (Hotel 1)\\ingredient_masterlist_birchsreet.xlsx'
# paths for checkscm
beverage_checkscm_path = 'data\\20231207_client_data\\export from Purchasing software\\CHECKSCM (Hotel 3)\\beverage_ingredient_masterlist_checkscm.xlsx'
ingredient_checkscm_path = 'data\\20231207_client_data\\export from Purchasing software\\CHECKSCM (Hotel 3)\\food_ingredient_masterlist_checkscm.xlsx'
recipe_checkscm_path = 'data\\20231207_client_data\\export from Purchasing software\\CHECKSCM (Hotel 3)\\food_recipe_checkscm.xlsx'
# paths for iscala
file_iscala_paths = ['data\\20231207_client_data\\export from Purchasing software\\ISCALA (Hotel 2)\\food_recipe_iscala.xlsx', 
              'data\\20231207_client_data\\export from Purchasing software\ISCALA (Hotel 2)\\beverage_recipe_iscala.xlsx',
                'data\\20231207_client_data\\export from Purchasing software\\ISCALA (Hotel 2)\\conversion_recipe_iscala.xlsx']
#path for manual data
manual_excel_files = [
    "data\\20231207_client_data\\recipe manual input on excel\\Hotel 1\\chef_manual_recipe_conversion_hotel1.xlsx",
    "data\\20231207_client_data\\recipe manual input on excel\\Hotel 1\\chef_manual_recipe_pizza_hotel1.xlsx",
]

# Output paths for the JSON files
# output paths for birchstreet
output_json_birchstreet_beverage_path = 'data_converted\\birchstreet\\beverage_recipe_birchstreet.json'
output_json_birchstreet_ingredient_path = 'data_converted\\birchstreet\ingredient_masterlist_birchstreet.json'
# output paths for checkscm
output_json_checkscm_beverage_path = 'data_converted\\checkscm\\beverage_ingredient_masterlist_checkscm.json'
output_json_checkscm_ingredient_path = 'data_converted\\checkscm\\food_ingredient_masterlist_checkscm.json'
output_json_checkscm_recipe_path = 'data_converted\\checkscm\\food_recipe_checkscm.json'
# output paths for iscala
output_file__iscala_paths = ["data_converted\\iscala\\food_recipe_iscala.json", 
                             "data_converted\\iscala\\beverage_recipe_iscala.json", 
                             "data_converted\\iscala\\conversion_recipe_iscala.json"]


# Identified header row indices for both files
header_row_index_beverage = 5  
header_row_index_ingredient = 0  

# Calling the function to convert both Excel files to CSV
convert_birchstreet_excel_to_json(beverage_birchstreet_path, ingredient_birchstreet_path, 
                          output_json_birchstreet_beverage_path, output_json_birchstreet_ingredient_path, 
                          header_row_index_beverage, header_row_index_ingredient)

convert_checkscm_excel_to_json(beverage_checkscm_path,ingredient_checkscm_path,recipe_checkscm_path,
                              output_json_checkscm_beverage_path,output_json_checkscm_ingredient_path,output_json_checkscm_recipe_path)

convert_iscala_excel_to_json(file_iscala_paths,output_file__iscala_paths)

for file_path in manual_excel_files:
    # Définir un nom de fichier JSON de sortie unique pour chaque fichier Excel
    json_output_path = file_path.replace('.xlsx', '_data.json')
    # Appeler la fonction d'extraction et de sauvegarde des données
    convert_manual_to_json(file_path, json_output_path)