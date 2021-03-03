#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#pip install xlwings


# In[13]:


#import xlwings as xw
#df = pd.read_excel('InputOutput.xlsm')
#xw.view(df)


# In[14]:


# viewing the dataframe via Xlwings directly in Excel can be very supportive, especially if you want to easily check lots of calculations
#df['Total']=(df['Column A']+df['Column B'] + df['Column C']) 
#xw.view(df)   


# In[15]:


# Make sure that this Jupyter Notebook is working. Make sure, that the Excel file is located in the same path as this file

import pandas as pd
import xlwings as xw
def main():
    wb = xw.Book.caller()
    sht = wb.sheets['Input']
    df = sht.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
    counter = []
# add Counter starting from one until max counter is reached 
    for index, row in df.iterrows():
        for x in range(int(row['Counter'])):
            counter.append(x+1)
# duplicate the rows according to Item specific max counter
    df = df.loc[df.index.repeat(df.Counter)]
# Add the counter number per row
    df['CounterDuplRow'] = counter
    df['Total']=(df['Column A']+df['Column B'] + df['Column C']) *df['CounterDuplRow']
    df2 = df.sort_values("Total").groupby("Item", as_index=False).last()
    shtout = wb.sheets['Output']
    shtout.range('a1').options(pd.DataFrame, index=False).value = df2
@xw.func
def hello(name):
    return f"Hello {name}!"
if __name__ == "__main__":
    xw.Book("InputOutput.xlsm").set_mock_caller()
    main()
# in case this Jupyter Notebook is working well, you can download it as a Python file (.py)
# to do so, just click on "file" in the menu above and choose "download as py"


# In[ ]:




