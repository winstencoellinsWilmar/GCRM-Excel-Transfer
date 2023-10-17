#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[2]:


# This function is to get DataFrame of distinct VAS Business Item from
# the Local VAS in Global CRM
# @returns DataFrame of distinct VAS
def get_vas_unique_business_item(vas_file_path):
    vas_excel = pd.read_excel(vas_file_path)
    vas_df = vas_excel.drop_duplicates(subset=["Business Item"]).reset_index(drop=True)

    vas_df = pd.DataFrame(vas_df, columns=["Business Item"])
    
    return vas_df


# In[3]:


# This function filters the VAS DataFrame record (row) based on
# the string
# @params: s -> string, df -> DataFrame
# @returns filtered DataFrame based on "s"
def get_vas_by_filter(s, df):
    return df[df["Business Item"].str.contains(s)]


# In[4]:


# This function is to drop the record of inactive row
# where status of Is Active is False
# @returns DataFrame
def drop_inactive(path):
    vas_df = pd.read_excel(path)

    return vas_df[vas_df["Is Active"] == True]


# In[5]:


# This function generates an excel file where business item
# is a REL_ (DEV_VAS_Relation sheet)
def generate_vas_rel_sheet(vas_df, unique_vas_bi_df):
    new_df = pd.DataFrame()

    vas_rel_temp = get_vas_by_filter("REL_", unique_vas_bi_df).to_numpy()

    for business_item in vas_rel_temp:
        desc_cols = []

        temp_df = vas_df[vas_df["Business Item"].str.contains(business_item[0])]
        temp_df = temp_df[["Business Item", "Key", "Value"]]

        for i in temp_df["Business Item"].to_numpy():
            s = i.split("_")
            desc_cols.append(f"Relation between {s[1]} & {s[2]}")
        
        obj_a_cols = [i.split("(")[0] for i in temp_df["Key"].to_numpy()]
        obj_b_cols = [i for i in temp_df["Value"].to_numpy()]

        temp_df.insert(loc=1, column="Description", value=desc_cols)
        temp_df.insert(loc=2, column="Object A Name", value=obj_a_cols)
        temp_df.insert(loc=3, column="Object B Name", value=obj_b_cols)

        new_df = pd.concat([new_df, temp_df], ignore_index=True)

    with pd.ExcelWriter('updated/DEV_VAS_Relation.xlsx') as writer:
        new_df.to_excel(writer)


# In[6]:


# This function generates an excel file where business item
# includes "Action Rights" string
def generate_action_rights(vas_df, unique_vas_bi_df):
    vas_action_rights = get_vas_by_filter("ActionRights", unique_vas_bi_df).to_numpy()

    for business_item in vas_action_rights:
        temp_df = vas_df[vas_df["Business Item"] == business_item[0]]
        temp_df = temp_df[["Business Item", "Key", "Value", "Country", "Trader Group"]]

        desc_cols = ['Create and Set Amendment Action Status' for i in range(len(temp_df))]

        temp_df.insert(loc=1, column="Description", value=desc_cols)

        s = "updated/DEV_VAS_" + business_item[0] + '.xlsx'

        with pd.ExcelWriter(s) as writer:
            temp_df.to_excel(writer)


# In[7]:


# This function generates an excel file where business item
# is STDCostSummaryPackage
def generate_std_cost_summary(vas_df, unique_vas_bi_df):
    vas_std_cost_summary = get_vas_by_filter("STDCostSummary", unique_vas_bi_df).to_numpy()

    temp_df = vas_df[vas_df["Business Item"].str.contains(vas_std_cost_summary[0][0])]
    temp_df = temp_df[["Business Item", "Key", "Value", "Country", "Comment"]]

    idn_std_cost_df = temp_df[temp_df["Country"] == "Indonesia"]
    sgp_std_cost_df = temp_df[temp_df["Country"] == "Singapore"]

    sgp_desc_cols = ["Cost summary calculation package for standard product" for desc in range(len(sgp_std_cost_df))]
    idn_desc_cols = ["Cost summary calculation package for standard product" for desc in range(len(idn_std_cost_df))]
    
    sgp_std_cost_df.insert(loc=1, column="Description", value=sgp_desc_cols)
    idn_std_cost_df.insert(loc=1, column="Description", value=idn_desc_cols)
    
    with pd.ExcelWriter('updated/DEV_VAS_STDCostSummaryPackage.xlsx') as writer:
        sgp_std_cost_df.to_excel(writer, sheet_name="SGP")
        idn_std_cost_df.to_excel(writer, sheet_name="IDN")


# In[8]:


# This function generates an excel file where business item
# is Port
def generate_port(vas_df, unique_vas_bi_df):
    vas_port = get_vas_by_filter("Port", unique_vas_bi_df).to_numpy()

    temp_df = vas_df[vas_df["Business Item"] == "Port"]
    temp_df = temp_df[["Business Item", "Key", "Value", "Country"]]

    port_desc = ["Dropdown for Port" for i in range(len(temp_df))]

    temp_df.insert(loc=1, column="Description", value=port_desc)

    with pd.ExcelWriter('updated/DEV_VAS_Port.xlsx') as writer:
        temp_df.to_excel(writer)


# In[9]:


# This function generates an excel file where business item
# include "Amendment Contract Flow"
def generate_amendment_contract_flow(vas_df, unique_vas_bi_df):
    vas_amendment = get_vas_by_filter("AmendmentContractFlow", unique_vas_bi_df).to_numpy() # AmendmentContractFlow & SupplierAmendmentContractFlow

    for business_item in vas_amendment:
        temp_df = vas_df[vas_df["Business Item"] == business_item[0]]
        temp_df = temp_df[["Business Item", "Key", "Value", "Status", "Country", "Comment"]]

        if business_item[0].__contains__("Supplier"):
            desc_cols = ["Workflow for SupplierAmendmentContract" for i in range(len(temp_df))]
        else:
            desc_cols = ["Workflow for AmendmenetContract" for i in range(len(temp_df))]

        temp_df.insert(loc=1, column="Description", value=desc_cols)

        s = "updated/DEV_VAS_" + business_item[0] + '.xlsx'

        with pd.ExcelWriter(s) as writer:
            temp_df.to_excel(writer)


# In[10]:


# This function generates an excel file where business item
# is STDCostSummaryPackage
def generate_sn_cost_summary(vas_df, unique_vas_bi_df):
    vas_sn_cost_summary = get_vas_by_filter("SNCostSummary", unique_vas_bi_df).to_numpy()

    temp_df = vas_df[vas_df["Business Item"] == vas_sn_cost_summary[0][0]]
    temp_df = temp_df[["Business Item", "Key", "Value", "Country", "Comment"]]

    desc_cols = ["Cost summary calculation package for standard product" for i in range(len(temp_df))]

    temp_df.insert(loc=1, column="Description", value=desc_cols)

    temp_df = temp_df.sort_values("Comment")
    
    with pd.ExcelWriter('updated/DEV_VAS_SN_CostSummaryPackage.xlsx') as writer:
        temp_df.to_excel(writer)


# In[11]:


# This function generates an excel file where business item
# is SN9015CostSummary
def generate_sn_9015(vas_df, unique_vas_bi_df):
    vas_sn_9015 = get_vas_by_filter("SN9015CostSummaryPackage", unique_vas_bi_df).to_numpy()

    temp_df = vas_df[vas_df["Business Item"] == vas_sn_9015[0][0]]
    temp_df = temp_df[["Business Item", "Key", "Value", "Country", "Comment"]]

    temp_df = temp_df.sort_values("Comment")

    desc_cols = ['' for i in range(len(temp_df))]

    temp_df.insert(loc=1, column="Description", value=desc_cols)

    with pd.ExcelWriter('updated/DEV_VAS_SN9015CostSummaryPackage.xlsx') as writer:
        temp_df.to_excel(writer)


# In[12]:


# This function generates an excel file where business item
# is SFCostSummary
def generate_sf_cost_summary(vas_df, unique_vas_bi_df):
    vas_sf_cost_summary = get_vas_by_filter("SF", unique_vas_bi_df).to_numpy()

    temp_df = vas_df[vas_df["Business Item"] == vas_sf_cost_summary[0][0]]
    temp_df = temp_df[["Business Item", "Key", "Value", "Country", "Comment"]]

    temp_df = temp_df.sort_values("Comment")

    desc_cols = ['' for i in range(len(temp_df))]

    temp_df.insert(loc=1, column="Description", value=desc_cols)

    with pd.ExcelWriter('updated/DEV_VAS_SFCostSummaryPackage.xlsx') as writer:
        temp_df.to_excel(writer)


# In[13]:


# This function generates an excel file where business item
# include Generate Quotation Flow
def generate_quote_flow(vas_df, unique_vas_bi_df):
    vas_quote_flow = get_vas_by_filter("QuoteFlow", unique_vas_bi_df).to_numpy()

    for business_item in vas_quote_flow:
        temp_df = vas_df[vas_df["Business Item"] == business_item[0]]
        temp_df = temp_df[["Business Item", "Key", "Value", "Country", "Comment"]]
        desc_cols = ['Get ' + business_item[0] for i in range(len(temp_df))]

        temp_df.sort_values("Key")
    
        temp_df.insert(loc=1, column="Description", value=desc_cols)
    
        s = 'updated/DEV_VAS_' + business_item[0] + '.xlsx'
        
        with pd.ExcelWriter(s) as writer:
            temp_df.to_excel(writer)


# In[14]:


# This function generates an excel file where business item
# is Quotation Packing
def generate_quotation_packing(vas_df, unique_vas_bi_df):
    vas_packing = get_vas_by_filter("Quotation_Packing", unique_vas_bi_df).to_numpy()

    temp_df = vas_df[vas_df["Business Item"] == vas_packing[0][0]]
    temp_df = temp_df[["Business Item", "Key", "Value", "Country"]]

    packing_desc = ["Dropdown for Packing" for i in range(len(temp_df))]

    temp_df.insert(loc=1, column="Description", value=packing_desc)

    with pd.ExcelWriter('updated/DEV_VAS_Quotation_Packing.xlsx') as writer:
        temp_df.to_excel(writer)


# In[15]:


def exclude_rows(path, filter):
    bi_arr = []
    distinct_vas_bi = get_vas_unique_business_item(path)

    for bi in filter:
        temp_arr = get_vas_by_filter(bi, distinct_vas_bi).to_numpy()
        
        if bi == "Port":
            temp_arr = [["Port"]]
            
        for item in temp_arr:
            bi_arr.append(item[0])

    # bi_arr.sort()
    # print(bi_arr)
    
    vas_df = pd.read_excel(path, index_col="Business Item")
    vas_df = vas_df.drop(bi_arr).reset_index()

    arr = vas_df["Business Item"].to_numpy()
        
    distinct = set(arr)

    return list(distinct)


# In[16]:


# This function is to generate DEV_VAS_Code sheet in the excel
def generate_code(vas_path, dev_vas_path, filter):
    new_df = pd.DataFrame()
    
    vas_df = pd.read_excel(vas_path)
    dev_vas_code_df = pd.read_excel(dev_vas_path)

    lst = exclude_rows(vas_path, filter)
    lst.sort()

    filtered_dev_vas_code_df = dev_vas_code_df[["Business Item", "Description"]]
    filtered_vas_df = vas_df[["Business Item", "Key", "Value", "Country", "Trader Group"]]

    unique_dev_vas_lst = get_vas_unique_business_item(dev_vas_path).to_numpy()
    to_arr = [i[0] for i in unique_dev_vas_lst]

    for bi in lst:
        temp_vas_df = filtered_vas_df[filtered_vas_df["Business Item"] == bi]
        
        if bi in to_arr:
            desc = filtered_dev_vas_code_df[filtered_dev_vas_code_df["Business Item"] == bi]["Description"].reset_index(drop=True)
            desc_cols = [desc[0] for i in range(len(temp_vas_df))]
        else:
            desc_cols = ['' for i in range(len(temp_vas_df))]

        temp_vas_df.insert(loc=1, column="Description", value=desc_cols)

        new_df = pd.concat([new_df, temp_vas_df], ignore_index=True)

    with pd.ExcelWriter('updated/DEV_VAS_Code.xlsx') as writer:
        new_df.to_excel(writer)


# In[17]:


# This is the main code that runs the program
if __name__ == "__main__":
    # Declare global var
    unique_vas_df = get_vas_unique_business_item('excel/vas.xlsx')
    vas_df = drop_inactive('excel/vas.xlsx')
    code_lst = ["REL_", "STDCostSummary", "Port", "AmendmentContractFlow", "SNCostSummary", "SN9015CostSummaryPackage", "QuoteFlow", 
                "Quotation_Packing", "SFCostSummary", "ActionRights"]

    # Generate all excel files
    generate_vas_rel_sheet(vas_df, unique_vas_df)
    generate_std_cost_summary(vas_df, unique_vas_df)
    generate_port(vas_df, unique_vas_df)
    generate_amendment_contract_flow(vas_df, unique_vas_df)
    generate_action_rights(vas_df, unique_vas_df)
    generate_sn_cost_summary(vas_df, unique_vas_df)
    generate_sn_9015(vas_df, unique_vas_df)
    generate_sf_cost_summary(vas_df, unique_vas_df)
    generate_quote_flow(vas_df, unique_vas_df)
    generate_quotation_packing(vas_df, unique_vas_df)
    generate_code('excel/vas.xlsx', 'excel/dev_vas_code.xlsx', code_lst)


# In[ ]:




