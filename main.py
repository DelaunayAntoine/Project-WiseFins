import pandas as pd 
import json

def convert_birchstreet_excel_to_json(file_path1, file_path2, output_file_path1, output_file_path2, header_row_index1, header_row_index2):
    """
    Converts a pair of Excel files to CSV files, considering the header row index for each file.

    :param file_path1: str, path to the first Excel file.
    :param file_path2: str, path to the second Excel file.
    :param output_file_path1: str, path where the first CSV file will be saved.
    :param output_file_path2: str, path where the second CSV file will be saved.
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
    Converts three Excel files to CSV files

    :param file_path1: str, path to the first Excel file.
    :param file_path2: str, path to the second Excel file.
    :param file_path3: str, path to the third Excel file.
    :param output_file_path1: str, path where the first json file will be saved.
    :param output_file_path2: str, path where the second json file will be saved.
    :param output_file_path3: str, path where the third jsonfile will be saved.
    """
    
    df1 = pd.read_excel(file_path1)
    df1.to_json(output_file_path1, orient="index")

    df2 = pd.read_excel(file_path2)
    df2.to_json(output_file_path2, orient="index")

    df3 = pd.read_excel(file_path3)
    df3.to_json(output_file_path3, orient="index")

    return f"Files converted and saved as '{output_file_path1}' ,'{output_file_path2}' and '{output_file_path3}"

def convert_iscala_excel_to_json(file_path1, file_path2,file_path3,output_file_path1,output_file_path2,output_file_path3):
    list_products: list[dict] = []
    current_product: dict = {}
    list_ingredients: list[dict] = []
    list_products2: list[dict] = []
    current_product2: dict = {}
    list_ingredients2: list[dict] = []
    list_products3: list[dict] = []
    current_product3: dict = {}
    list_ingredients3: list[dict] = []
    
    df: pd.DataFrame = pd.read_excel(file_path1,sheet_name='Sheet1',skiprows=6,header=None)
    df = df.dropna(axis = 0, how = 'all')
    df = df[~df[0].isin(['- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'])]
    df2: pd.DataFrame = pd.read_excel(file_path2,sheet_name='Sheet1',skiprows=6,header=None)
    df2 = df2.dropna(axis = 0, how = 'all')
    df2 = df2[~df2[0].isin(['- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'])]
    df3: pd.DataFrame = pd.read_excel(file_path2,sheet_name='Sheet1',skiprows=6,header=None)
    df3 = df3.dropna(axis = 0, how = 'all')
    df3 = df3[~df3[0].isin(['- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'])]
   
    for row in df.iterrows():
        # First item row
        if str(row[0]).find('Template : ') != -1:
            current_product = {
                "name": row[0][12:]
            }

            continue
        
        # Last item row
        if str(row[0]) == 'Total for Stock Item':
            current_product['total'] : dict = {
                # TODO rename
                'total_first_col': row[11],
                'total_second_col': row[17],
                'total_third_col': str(row[21]),
                'total_fourth_col': row[24]
                
            }

            current_product['ingredients'] = list_ingredients
            list_products.append(current_product)
            
            list_ingredients = []

            continue
    

        # Intermediate row (Ingredient)

        list_ingredients.append({
            'code': row[0],
            'name': row[8],
            # TODO rename those columns
            'unkown_col_1': row[14],
            'unkown_col_1': row[20],
            'unkown_col_1': row[24],
            'unkown_col_1': row[32],
            'unkown_col_1': row[36],
            'unkown_col_1': row[42]
        })

    with open(output_file_path1, "w") as outfile:
        # json_data refers to the above JSON
        json.dump(list_products, outfile, indent=4)


# Paths to the Excel files
# paths for birchstreet
beverage_birchstreet_path = 'data\\20231207_client_data\\export from Purchasing software\\BIRCHSTREET (Hotel 1)\\beverage_recipe_birchstreet.xlsx'
ingredient_birchstreet_path = 'data\\20231207_client_data\\export from Purchasing software\BIRCHSTREET (Hotel 1)\\ingredient_masterlist_birchsreet.xlsx'
# paths for checkscm
beverage_checkscm_path = 'data\\20231207_client_data\\export from Purchasing software\\CHECKSCM (Hotel 3)\\beverage_ingredient_masterlist_checkscm.xlsx'
ingredient_checkscm_path = 'data\\20231207_client_data\\export from Purchasing software\\CHECKSCM (Hotel 3)\\food_ingredient_masterlist_checkscm.xlsx'
recipe_checkscm_path = 'data\\20231207_client_data\\export from Purchasing software\\CHECKSCM (Hotel 3)\\food_recipe_checkscm.xlsx'
# paths for iscala
food_iscala_path = 'data\\20231207_client_data\\export from Purchasing software\\ISCALA (Hotel 2)\\food_recipe_iscala.xlsx'
beverage_iscala_path = 'data\\20231207_client_data\\export from Purchasing software\ISCALA (Hotel 2)\\beverage_recipe_iscala.xlsx'
conversion_iscala_path = 'data\\20231207_client_data\\export from Purchasing software\\ISCALA (Hotel 2)\\conversion_recipe_iscala.xlsx'


# Output paths for the JSON files
# output paths for birchstreet
output_json_birchstreet_beverage_path = 'data_converted\\birchstreet\\beverage_recipe_birchstreet.json'
output_json_birchstreet_ingredient_path = 'data_converted\\birchstreet\ingredient_masterlist_birchstreet.json'
# output paths for checkscm
output_json_checkscm_beverage_path = 'data_converted\\checkscm\\beverage_ingredient_masterlist_checkscm.json'
output_json_checkscm_ingredient_path = 'data_converted\\checkscm\\food_ingredient_masterlist_checkscm.json'
output_json_checkscm_recipe_path = 'data_converted\\checkscm\\food_recipe_checkscm.json'
# output paths for iscala
output_json_iscala_food_path ='data_converted\\iscala\\food_recipe_iscala.json'
output_json_conversion_iscala_path = 'data_converted\\iscala\\beverage_recipe_iscala.json'
output_json_beverage_iscala_path = 'data_converted\\iscala\\conversion_recipe_iscala.json'

# Identified header row indices for both files
header_row_index_beverage = 5  
header_row_index_ingredient = 0  

# Calling the function to convert both Excel files to CSV
# convert_birchstreet_excel_to_json(beverage_birchstreet_path, ingredient_birchstreet_path, 
#                           output_json_birchstreet_beverage_path, output_json_birchstreet_ingredient_path, 
#                           header_row_index_beverage, header_row_index_ingredient)
# convert_checkscm_excel_to_json(beverage_checkscm_path,ingredient_checkscm_path,recipe_checkscm_path,
#                               output_json_checkscm_beverage_path,output_json_checkscm_ingredient_path,output_json_checkscm_recipe_path)

convert_iscala_excel_to_json(food_iscala_path,output_json_iscala_food_path)
