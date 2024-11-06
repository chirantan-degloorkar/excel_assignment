import pandas as pd
from core import logger
import json

with open('config.json', 'r') as f:
    config = json.load(f)
    
sheet_name = config['sheet_name']

def find_table_in_excel(df):
    """Find the table start in the excel file

    Args:
        df (pd.DataFrame): Dataframe from the excel file

    Returns:
        pd.DataFrame: Dataframe with only the table
    """
    row_num = 0
    try:
        for i, row in df.iterrows():
            if row.apply(lambda x: isinstance(x, str)).sum() > 50:
                # df = pd.read_excel('input/input1.xlsx',sheet_name='3a. Bid Sheet- 100% volume', skiprows=i+1)
                df = pd.DataFrame(columns=df.iloc[i], data=df.values[i+1:])
                # check if first column of the table has numerical values
                df = df.loc[df.iloc[:,1].apply(lambda x: isinstance(x, (int, float)))]               
                
                row_num = i
                # df.dropna(inplace=True)
        return df, row_num
    except Exception as e:
        # print(e)
        logger.log_message(message=f"ERROR: Finding table in Excel", level=1)
        raise e

def is_matching(df, template):
    """ Check if the columns of the dataframe match the template

    Args:
        df (pd.DataFrame): Dataframe to be checked
        template (pd.DataFrame): Template to be checked against

    Returns:
        bool: True if the columns match, False otherwise
    """
    try:
        return len(template.columns) == len(df.columns) and template.columns.equals(df.columns)
    except Exception as e:
        # print(e)
        logger.log_message(message=f"ERROR: Checking if dataframes match", level=1)
        raise e
    
def concatenate_dataframes(dataframes, template_path):
    """ Concatenate dataframes if they match

    Args:
        dataframes (list): List of dataframes to be concatenated

    Returns:
        pd.DataFrame: Concatenated dataframe
    """
    try:
        template = pd.read_excel(template_path, sheet_name=sheet_name, engine='openpyxl')
        template, _ = find_table_in_excel(template)
        
        not_matching_dfs = []
        
        for df in dataframes:
            if is_matching(df, template):
                template = pd.concat([template, df], axis=0)
            else:
                not_matching_dfs.append(df)
                continue

        if 'abcPart No.' in template.columns:
            template = template[template['abcPart No.'].apply(lambda x: isinstance(x, int))]
        
        print(template.shape)
            
        return template
    
    except Exception as e:
        print(e)
        logger.log_message(message=f"ERROR: Combining Excels", level=1)
        

def format_and_save(result_df):
    try:
        df = pd.read_excel('input/Consolidation_Assignment.xlsx', sheet_name=sheet_name, engine='openpyxl')
        
        _, row_num = find_table_in_excel(df)
        location = 'output/Consolidation_Assignment.xlsx'
        start_row = len(df) 
        
        with pd.ExcelWriter(location, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            result_df.to_excel(writer, sheet_name=sheet_name, startrow=row_num+3, header=False, index=False)
            
        logger.log_message(message=f"EXEC: Consolidated file saved at {location}", level=0)
    except Exception as e:
        # print(e)
        logger.log_message(message=f"ERROR: In saving Consolidated file", level=1)
        raise e

