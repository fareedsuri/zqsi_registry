import os
import sys
# print(sys.executable)
from openpyxl import load_workbook
import pandas as pd
from itertools import chain
import re


def get_var_list(folder):
    # print(folder)
    sys.setrecursionlimit(999)
    to_exclude = 'load,rename,recode,refilter,df,title'.split(',')

    # Get the current directory
    # current_directory = os.getcwd()

    # Define the name of the Excel file you want to open
    file_name = "recode.xlsx"

    # Construct the full path to the Excel file
    file_path = os.path.join(folder, file_name)

    df_original_ = pd.read_excel(file_path, sheet_name="original")
    df_recode_ = pd.read_excel(file_path, sheet_name="recode")
    df_analysis_ = pd.read_excel(file_path, sheet_name="analysis")

    df_recode = pd.concat([df_original_, df_recode_],
                          axis=0).dropna(subset=['nvar'])
    df_analysis = df_analysis_[(df_analysis_['include'] == 1) & (
        ~df_analysis_['command'].isin(to_exclude))]
    # col_2_vars = df_analysis[[2]][2].str.extract(r'var:(.+),')[0].unique()
    col_2_vars_ = []
    if not df_analysis[[2]][2].dropna().empty:
        col_2_vars_1 = df_analysis[[2]][2].str.findall(
            r'[#\$]([\w_]+)').dropna()
        col_2_vars_2 = df_analysis[[2]][2].str.findall(r'var:(.+),').dropna()
        col_2_vars_ = col_2_vars_1 + col_2_vars_2
    # set to get unique, chain to flatten
    col_2_vars = set(list(chain.from_iterable(col_2_vars_)))

    col_1_vars_ = []
    if not df_analysis[[1]][1].dropna().empty:
        col_1_vars_ = df_analysis[[1]][1].str.findall(
            r'[#\$]([\w_]+)').dropna()
    # set to get unique, chain to flatten
    col_1_vars = set(list(chain.from_iterable(col_1_vars_)))

    vars_list = set(list(col_2_vars) + list(col_1_vars))
    if not vars_list:
        return []
    vars_list = sorted([x for x in vars_list if pd.notna(x)])
    df_var_list = pd.DataFrame(vars_list, columns=['var'])
    df_var_list['definition'] = ''

    new_fx = []

    def replace_pattern(match):
        # pattern = match.group(0)  # Get the matched pattern
        num = match.group(1) if match.group(1) is not None else match.group(2)
        # print(num)
        # Replace with corresponding 'nvar' value
        return "$" + df_recode.loc[df_recode['id'] == int(num), 'nvar'].iloc[0]

    def iterative_build_fx(chain, nvar, built, level):
        chain.append(nvar)
        # get formula for the nvar
        try:
            fx = df_recode[df_recode['nvar'] == nvar]['new_fx'].values[0]
        except IndexError:
            raise ValueError(f"Error: 'nvar' {nvar} not found in 'df_recode'")
        var = df_recode[df_recode['nvar'] == nvar]['var'].values[0]
        if pd.isna(fx):
            built = built + '\t' * level + nvar + '=' + var + '\n'
            chain.pop()
            return built
        fx_str = str(fx) if not isinstance(fx, str) else fx
        # add formula to built using level as tabs
        built = built + '\t' * level + nvar + '=' + fx_str + '\n'
        # get list of each nvar in the formula
        nvars = [x for x in set(re.findall(r'[#\$]([\w_]+)', fx))]
        # and for each nvar in the list, call iterative_build_fx
        for nvar_ in nvars:
            if nvar_ in chain:
                raise ValueError(f'Circular reference ', nvar_, ' in ', chain)
            built = iterative_build_fx(chain, nvar_, built, level+1)
        chain.pop()
        return built

    for value in df_recode['fx']:
        if not pd.isna(value):
            new_fx.append(re.sub(r'[#\$](\d+)', replace_pattern, value))
        else:
            new_fx.append(value)
    df_recode['new_fx'] = new_fx
    for nvar in vars_list:
        # print(nvar)
        df_var_list.loc[df_var_list['var'] == nvar,
                        'definition'] = iterative_build_fx([], nvar, '', 0)
    return df_var_list


# print(get_var_list('E:/SynologyDrive/Data/CREST/crest_analysis'))

# print(get_var_list('/Users/fareedsuri/SynologyDrive/Data/CREST/crest_analysis'))
