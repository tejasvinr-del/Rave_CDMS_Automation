import numpy as np
import re
def replace_trans(df):
    """
    Standardizes and re-sequences the Transpose column.
    Ensures sets of variables from the same Source Table share the same Transpose number,
    but different Source Tables get uniquely incremented numbers.
    """
    df['Transpose'] = df['Transpose'].fillna('')
    
    tmp1 = df[['Target Table', 'Source Table', 'Transpose']]
    
    # 1. Isolate simple transpose rules (e.g., CE1, DA1)
    tmp2 = tmp1.loc[
        (tmp1['Transpose'] != '') & 
        (tmp1['Transpose'].str.upper() != 'ALL') & 
        (~tmp1['Transpose'].str.contains(',|-', regex=True))
    ]
    
    tmp3 = tmp2.drop_duplicates(subset=['Target Table', 'Source Table', 'Transpose']).copy()
    tmp3['Transpose Num'] = tmp3['Transpose'].str.extract(r'(\d+)').fillna(0).astype(int)
    tmp3['Transpose Char'] = tmp3['Transpose'].str.extract(r'(\D+)')
    
    srt_list = ['Target Table', 'Transpose Char', 'Source Table']
    tmp4 = tmp3.sort_values(by=srt_list).copy()
    
    tmp4['seq'] = tmp4.groupby(['Target Table', 'Transpose Char']).cumcount() + 1
    tmp4['Transpose_1'] = tmp4['Transpose Char'] + tmp4['seq'].astype(str)
    
    tmp5 = tmp4[['Target Table', 'Source Table', 'Transpose', 'Transpose_1']]
    tmp6 = df.merge(tmp5, on=['Target Table', 'Source Table', 'Transpose'], how="left")
    
    # 2. Handle complex rules (e.g., CE1,CE2 or CE1-CE3)
    df_other = tmp1.loc[tmp1['Transpose'].str.contains(',|-', regex=True)]
    df_other_1 = df_other.drop_duplicates(subset=['Target Table', 'Source Table', 'Transpose']).copy()
    
    table_mappings = {}
    for _, row in tmp5.iterrows():
        tt = row['Target Table']
        old_t = row['Transpose']
        new_t = row['Transpose_1']
        
        if old_t != new_t:
            if tt not in table_mappings:
                table_mappings[tt] = {}
            table_mappings[tt][old_t] = new_t
            
    def safe_replace(row):
        val = row['Transpose']
        tt = row['Target Table']
        
        if tt in table_mappings and isinstance(val, str):
            return re.sub(r'[A-Za-z]+\d+', lambda m: table_mappings[tt].get(m.group(0), m.group(0)), val)
        return val
        
    df_other_1['Transpose_2'] = df_other_1.apply(safe_replace, axis=1)
    
    tmp7 = tmp6.merge(
        df_other_1[['Target Table', 'Source Table', 'Transpose', 'Transpose_2']], 
        on=['Target Table', 'Source Table', 'Transpose'], 
        how="left"
    )
    
    # 3. Apply the correct mapped column based on condition
    tmp7["Transpose"] = np.select(
        [
            (tmp7["Transpose_1"].notna()) & (tmp7['Transpose'] != tmp7['Transpose_1']),
            (tmp7["Transpose_2"].notna()) & (tmp7['Transpose'] != tmp7['Transpose_2'])
        ],
        [
            tmp7["Transpose_1"],
            tmp7["Transpose_2"],
        ],
        default=tmp7["Transpose"],
    )
    
    tmp7 = tmp7.drop(columns=['Transpose_1', 'Transpose_2'], errors='ignore')
    tmp7['Transpose'] = tmp7['Transpose'].replace('', None)
    
    return tmp7
