import pandas as pd
import sys

def transform(file):
    sales_a=pd.read_excel(file)
    size=sales_a.shape[0]
    for i in range(size):
        if str(sales_a.iloc[i,0]).startswith("** END OF REPORT **"):
            size=i
            break

    i=0
    data = pd.DataFrame()
    while i<size:
        if str(sales_a.iloc[i,0]).startswith("Category: "):
            header=sales_a.iloc[i+1, :]
            for j in range(i+2,size):
                if str(sales_a.iloc[j,0]).startswith("Sub Total:"):
                    cat=sales_a.iloc[i]
                    cat = cat[0].split("Category: ")
                    parent_grp=cat[1]
                    sales_data = sales_a.iloc[i+2:j, :].copy()
                    sales_data.loc[:, 'recipeGroupParent'] = parent_grp
                    data= pd.concat([data, sales_data], axis=0)
                    i=j+3
        else:
            i+=1
    data.replace("", float("NaN"), inplace=True) 
    data.dropna(how='all', axis=1, inplace=True)
    header=header.dropna()
    header_list=header.tolist()
    header_list.append('recipeGroupParent')
    data.columns = header_list
    data.to_excel("output_menu_sales_analysis_transformed.xlsx", index=False)
if __name__ == "__main__":
    file = sys.argv[1]
    transform(file)
    print(f"Data processed and saved to output_menu_sales_analysis_transformed")