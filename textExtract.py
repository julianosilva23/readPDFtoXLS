
# coding: utf-8

# In[1]:


import numpy as np
import docx2txt
import glob
import pandas as pd


# In[24]:


# extract text
def ExtrairTextDoc(file):
    text = docx2txt.process(file)
    text = text.replace('\n', ',').replace('\t', ',').split(',')
    text = list(filter(None, text))
    return [text, file]


# In[76]:


valores = []
path = []
files = glob.glob('./fichas1/*.docx')
for i in files:
    row = ExtrairTextDoc(i)
    valores.append(row[0])
    path.append(row[1][10:])


# In[97]:


df = pd.DataFrame(valores)
df['disciplina'] = path
cols = list(df.columns.values)
cols.sort(key = lambda item: ([str,int].index(type(item)), item))
df = df[cols]
df = df.dropna(axis=1, how='all')


# In[98]:


writer = pd.ExcelWriter('resultado_fichas1.xlsx')
df.to_excel(writer,'Sheet1')
writer.save()

