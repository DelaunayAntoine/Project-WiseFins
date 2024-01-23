import pandas as pd
import numpy as np
def extract_first_8_chars(recipe):
    return recipe[:8]

def main():



    b_r_birch=pd.read_excel('beverage_recipe_birchstreet.xlsx',header=5)
    b_r_birch = b_r_birch[b_r_birch['Recipe type'] == 'Recipe']

    b_i_birch=pd.read_excel('ingredient_masterlist_birchsreet.xlsx')

    c_i_birch=pd.read_excel('Category_ingredient_BirchStreet.xlsx')
    agbsrc=pd.read_excel('agbcode_file.xlsx')

    # Define additional columns
    additional_columns = {
        'resto': 'birchstreet',
        'recipeGroup': 'Beverage',
        'recipeGroupParent': np.nan,
        'agbcode': np.nan,
        'category': np.nan,
        'subcategory': np.nan,
        'photo': 'no',
        'cost': np.nan,
        'InventoryCode': np.nan,
        'subRecipe': np.nan,
        'portion': np.nan,
        'statut': np.nan,
        'selling': np.nan
    }

    # Apply additional columns to Birch Street beverage recipe data
    b_r_birch = b_r_birch.assign(**additional_columns)

    # Rename columns for consistency
    column_rename_mapping = {
        'Ingredient/Subrecipe name': 'recipeCompose',
        'Dish/Recipe name': 'recipeName',
        'UOM': 'unit'
    }
    b_r_birch.rename(columns=column_rename_mapping, inplace=True)

    # Copy the 'Quantity' column to 'quantityAfter'
    b_r_birch['quantityAfter'] = b_r_birch['Quantity'].copy()

    # Apply the function to create a new 'SKU' column
    b_r_birch['SKU'] = b_r_birch['recipeCompose'].apply(extract_first_8_chars)

    # Merge with Birch Street ingredient data to get 'Unit price'
    merged_df = pd.merge(b_r_birch, b_i_birch[['Supplier SKU', 'Unit price']], how='left', left_on='SKU', right_on='Supplier SKU')
    b_r_birch['cost'] = merged_df['Unit price']

    # Merge with Birch Street category data to get 'AGBCode'
    merged_df = pd.merge(b_r_birch, c_i_birch[['AGBCode', 'Part #']], how='left', left_on='SKU', right_on='Part #')
    b_r_birch['agbcode'] = merged_df['AGBCode']

    # Merge with AGB code file to get category and subcategory information
    merged_df = pd.merge(b_r_birch, agbsrc[['Code\nAGB', 'WiseFins EN Category', 'WiseFins EN Subcategory']], how='left', left_on='agbcode', right_on='Code\nAGB')
    b_r_birch['category'] = merged_df['WiseFins EN Category']
    b_r_birch['subcategory'] = merged_df['WiseFins EN Subcategory']

    # Define column order
    column_order = ['resto', 'subRecipe', 'recipeGroup', 'recipeGroupParent', 'recipeName', 'recipeCompose', 'agbcode', 'Quantity', 'cost', 'InventoryCode', 'unit', 'quantityAfter', 'portion', 'statut', 'photo', 'selling', 'category', 'subcategory']

    # Reorder DataFrame columns
    b_r_birch = b_r_birch[column_order]

    # Save the DataFrame to an Excel file
    b_r_birch.to_excel("output_b_r_birch.xlsx", index=False)

if __name__ == "__main__":
    main()
