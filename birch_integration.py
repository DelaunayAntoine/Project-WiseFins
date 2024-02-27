import pandas as pd
import numpy as np

def extract_first_8_chars(recipe):
    return recipe[:8]
def extract_last_8_chars(recipe):
    return recipe[-8:]
def main():

    b_r_birch=pd.read_excel('beverage_recipe_birchstreet.xlsx',header=5)
    b_r_birch = b_r_birch[b_r_birch['Recipe type'] == 'Recipe']

    b_i_birch=pd.read_excel('ingredient_masterlist_birchsreet.xlsx')

    c_i_birch=pd.read_excel('Category_ingredient_BirchStreet.xlsx')
    agbsrc=pd.read_excel('agbcode_file.xlsx')
    sales_a=pd.read_excel('output_menu_sales_analysis_transformed.xlsx')

    agbsrc=agbsrc[['Code AGB', 'WiseFins FR Category','WiseFins EN Category','WiseFins FR Subcategory','WiseFins EN Subcategory']]

    agbsrc.drop_duplicates(inplace=True)

    agbsrc['Code AGB']=agbsrc['Code AGB'].astype(str)
    # Define additional columns
    additional_columns = {
        'resto': 'birchstreet',
        'recipeGroup': 'Beverage',
        'subRecipe': np.nan,
        'statut' : 'ACTIVE',
        'photo' : 'no',
        'portion' : 1
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
    b_r_birch['SKU'] = b_r_birch['SKU'].astype(str)
    b_i_birch['Supplier SKU']=b_i_birch['Supplier SKU'].astype(str)
    # Calculate the average price for each unique code
    b_i_birch = b_i_birch.groupby('Supplier SKU', as_index=False)['Unit price'].mean()
    b_i_birch['Supplier SKU']=b_i_birch['Supplier SKU'].astype(str)

    # Merge with Birch Street ingredient data to get 'Unit price'
    b_r_birch = b_r_birch.merge(b_i_birch[['Supplier SKU', 'Unit price']], how='left', left_on='SKU', right_on='Supplier SKU')
    b_r_birch.drop_duplicates(inplace=True)
    b_r_birch.rename(columns={'Unit price': 'cost'}, inplace=True)
    c_i_birch=c_i_birch[['Part #','AGBCode']]
    c_i_birch.drop_duplicates(inplace=True)

    # Merge with Birch Street category data to get 'AGBCode'
    b_r_birch = b_r_birch.merge(c_i_birch[['AGBCode','Part #']], how='left', left_on='SKU', right_on='Part #')

    b_r_birch['AGBCode'].fillna(0, inplace=True)

    # Convert the column to integer
    b_r_birch['AGBCode'] = b_r_birch['AGBCode'].astype(str)
    agbsrc['Code AGB'] = agbsrc['Code AGB'].str.strip()
    b_r_birch['AGBCode'] = b_r_birch['AGBCode'].str.strip()

    # Merge with AGB code file to get category and subcategory information
    b_r_birch = b_r_birch.merge(agbsrc[['Code AGB','WiseFins EN Category','WiseFins EN Subcategory']], how='left', left_on='AGBCode', right_on='Code AGB')
    b_r_birch.rename(columns={'WiseFins EN Category': 'category','WiseFins EN Subcategory': 'subcategory'}, inplace=True)
    # Define column order
    column_order = ['resto', 'subRecipe', 'recipeGroup', 'recipeName', 'recipeCompose', 'AGBCode','Code AGB', 'Quantity', 'cost', 'unit', 'quantityAfter','statut', 'photo','category','subcategory','portion','SKU']
    # Reorder the columns
    b_r_birch = b_r_birch[column_order]
    sales_a['code_grpparent'] = sales_a['Code'].apply(extract_last_8_chars)
    b_r_birch['code_grp']=b_r_birch['recipeName'].apply(extract_last_8_chars)
    sales_a['code_grpparent']=sales_a['code_grpparent'].astype(str)
    b_r_birch['code_grp']=b_r_birch['code_grp'].astype(str)
    b_r_birch = b_r_birch.merge(sales_a[['Sales Price','recipeGroupParent','code_grpparent']], how='left', left_on='code_grp', right_on='code_grpparent')
    b_r_birch.rename(columns={'Sales Price': 'selling'}, inplace=True)
    b_r_birch['selling'] = b_r_birch['selling'].fillna(0.0001)
    b_r_birch.rename(columns={'SKU': 'InventoryCode'}, inplace=True)
    b_r_birch['recipeName']=b_r_birch['recipeName'].apply(lambda x: x[10:-8])
    b_r_birch['recipeName']=b_r_birch['recipeName'].str.strip()
    b_r_birch['AGBCode'] = b_r_birch['AGBCode'].astype(int)
    column_order = ['resto', 'subRecipe', 'recipeGroup', 'recipeGroupParent', 'recipeName', 'recipeCompose', 'AGBCode', 'Quantity', 'cost', 'InventoryCode', 'unit', 'quantityAfter','portion', 'statut', 'photo', 'selling', 'category', 'subcategory']
    # Reorder the columns
    b_r_birch = b_r_birch[column_order]
    b_r_birch.to_excel("output_birchstreet.xlsx", index=False)

if __name__ == "__main__":
    main()
