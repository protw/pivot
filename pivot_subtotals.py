''' Pivot table with all subtotals by rows and columns.

    Inspired by the article: "Tabulating Subtotals Dynamically in Python Pandas 
    Pivot Tables" by Will Keefe, Jun 24, 2022, https://medium.com/p/6efadbb79be2
    
    Author: @oleghbond, https://protw.github.io/oleghbond
    Date: April 13, 2023
'''

import pandas as pd, numpy as np, os

def pivot_w_subtot(df, values, indices, columns, aggfunc=np.nansum, 
                   fill_value=np.nan, margins=False):
    ''' Adds tabulated subtotals to pandas pivot tables with multiple 
        hierarchical indices.
        
        Args:
        - df - dataframe used in pivot table
        - values - values used to aggregrate
        - indices - ordered list of indices to aggregrate by
        - columns - columns to aggregrate by
        - aggfunc - function used to aggregrate (np.max, np.mean, 
          np.sum, etc)
        - fill_value - value used to in place of empty cells
        
        Returns:
        - flat table with data aggregrated and tabulated
    '''
    listOfTable = []
    for indexNumber in range(len(indices)):
        n = indexNumber + 1
        if n == 1:
            table = pd.pivot_table(df, values=values, index=indices[:n], 
                                   columns=columns, aggfunc=aggfunc, 
                                   fill_value=fill_value, margins=margins)
        else:
            table = pd.pivot_table(df, values=values, index=indices[:n], 
                                   columns=columns, aggfunc=aggfunc, 
                                   fill_value=fill_value)
        table = table.reset_index()
        for column in indices[n:]:
            table[column] = ''
        listOfTable.append(table)
    concatTable = pd.concat(listOfTable).sort_index()
    concatTable = concatTable.set_index(keys=indices)
    concatTable.sort_index(axis=0, ascending=True, inplace=True)

    return concatTable

def pivot_w_subtot2(df, values, indices, columns, aggfunc=np.nansum, 
                    fill_value=np.nan):

    ''' The problem with the pivot_w_subtot method is that it only builds a 
        vertical (by rows) pivot table nicely. That is, a general summary 
        (of the highest level) is not performed for all columns. Adding the 
        margins argument (see Will Keefe's original article) does not 
        completely solve this problem.
        
        Therefore, the purpose of pivot_w_subtot2 is to build, on the basis 
        of pivot_w_subtot, a summary table for all levels vertically (by rows) 
        and horizontally (by columns), including the overall summary of the 
        top level. 
        
        Args and Returns are similar to `pivot_w_subtot`. '''

    df_long = df.copy()
    df_long.fillna('-', inplace=True)
    
    ''' Creating two new temporary (dummy) columns to obtain total summaries 
        vertically (by rows) and horizontally (by columns) only at the top level. '''

    df_long['All_cols'] = 0 # For the summary by columns at uppermost level
    df_long['All_rows'] = 0 # For the summary by rows at uppermost level
    
    ''' First, we create a vertical summary, that is, for all rows, including the 
        total summary of the top level of the input table `df_long` (due to the 
        inclusion of the `All_cols` column and the corresponding index).  '''

    df_wide = pivot_w_subtot(df=df_long, values=values, 
                             indices=['All_rows',] + indices, 
                             columns=['All_cols',] + columns, 
                             aggfunc=aggfunc, 
                             fill_value=fill_value)

    ''' We expand the "wide" table `df_wide` again into the "long" `df_long`, but this 
        time it contains the summaries of all levels vertically. '''

    df_long = pd.melt(df_wide, value_name='value', ignore_index=False).reset_index()

    ''' Now from the long table `df_long` we create a summary horizontally (that is, 
        on all columns), but placing them vertically. This is achieved by assigning 
        all columns to the `indices` argument and, conversely, all rows to the 
        `columns` argument. And, at the end, we transpose (.T) the resulting table, 
        that is, we change the rows (axis=0) and columns (axis=1). '''

    df_wide = pivot_w_subtot(df=df_long, values='value', 
                             columns=['All_rows',]+indices, 
                             indices=['All_cols',]+columns, 
                             aggfunc=aggfunc, 
                             fill_value=fill_value).T

    ''' We modify the vertical and horizontal indices, for that we cut off their 
        temporary uppermost levels - `All_cols` and `All_rows`. '''

    df_wide.index = df_wide.index.droplevel(0)
    df_wide.columns = df_wide.columns.droplevel(0)

    return df_wide

if __name__ == '__main__':
    
    data_file = 'test_data/sampledatafoodsales.xlsx'

    df = pd.read_excel(data_file)
    
    val = 'TotalPrice'
    idx = ['Category', 'Product']
    cols = ['Region', 'City']

    ''' Standard Pivot '''
    pt0 = pd.pivot_table(df, values=val, index=idx, columns=cols, 
                         aggfunc=np.nansum)
    ''' Will Keefe's Pivot '''
    pt1 = pivot_w_subtot(df=df, values=val, indices=idx, columns=cols, 
                         aggfunc=np.nansum, margins=True)
    ''' My Pivot '''
    pt2 = pivot_w_subtot2(df=df, values=val, indices=idx, columns=cols, 
                          aggfunc=np.nansum)

    file_name, file_extension = os.path.splitext(data_file)
    result_file = file_name + '_pivtab.xlsx'

    with pd.ExcelWriter(result_file, engine='xlsxwriter') as writer:
        pt0.to_excel(writer, sheet_name="Standard Pivot")
        pt1.to_excel(writer, sheet_name="Will Keefe's Pivot")
        pt2.to_excel(writer, sheet_name="My Pivot")
