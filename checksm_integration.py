import pandas as pd
import numpy as np

def main():

    f_r_check=pd.read_excel('food_recipe_checkscm.xlsx')
    # Define additional columns
    additional_columns = {
        'resto': 'checkscm',
        'subRecipe': np.nan,
        'recipeCompose': np.nan,
        'agbcode': np.nan,
        'quantity': np.nan,
        'InventoryCode': np.nan,
        'quantityAfter': np.nan,
        'photo': np.nan,
        'category': np.nan,
        'subcategory': np.nan
    }

    # Apply additional columns
    f_r_check = f_r_check.assign(**additional_columns)
    column_rename_mapping = {
        'Class': 'recipeGroupParent',
        'Recipe Description': 'recipeName',
        'Unit Cost': 'cost',
        'Serve Size': 'unit',
        'No. of Serve': 'portion',
        'Status': 'statut',
        'Suggested Selling Price': 'selling'
    }
    f_r_check.rename(columns=column_rename_mapping, inplace=True)
    # Rename columns as needed
    f_r_check.rename(columns={'Recipe Group': 'recipeGroup', 'Class': 'recipeGroupParent', 'Recipe Description': 'recipeName', 'Unit Cost': 'cost', 'Serve Size': 'unit', 'No. of Serve': 'portion', 'Status': 'statut', 'Suggested Selling Price': 'selling'}, inplace=True)

    # Reorder columns
    column_order = ['resto', 'subRecipe', 'recipeGroup', 'recipeGroupParent', 'recipeName', 'recipeCompose', 'agbcode', 'quantity', 'cost', 'InventoryCode', 'unit', 'quantityAfter', 'portion', 'statut', 'photo', 'selling', 'category', 'subcategory']
    f_r_check = f_r_check[column_order]

    # Save the DataFrame to an Excel file
    f_r_check.to_excel("output_f_r_check.xlsx", index=False)


if __name__ == "__main__":
    main()