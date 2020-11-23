#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd   #Importing Pandas Module for data analysis.


# In[3]:


df_1 = pd.read_excel("/home/deep/Downloads/KPMG_VI_New_raw_data_update_final.xlsx", 1) #Reding data from Transaction sheet.


# In[4]:


df_1.head() #displaying only top 5 rows.


# In[5]:


for i in df_1.columns: #Running a loop to iterate through and each and every column values and want to see the occurrence of each column values so used the value_counts() function.
    print(f"{df_1[i].value_counts()}")
    print("-" * 50)
    


# In[6]:


print(df_1.shape) #Find out number of rows and columns we have
df_1.isnull().sum() #to check missing values in the dataframe.


# In[7]:


df_1.dropna(inplace=True)  #If we have any missing values for any column we have to remove them and for this used dropna() function.
print(df_1.shape)
df_1.drop_duplicates() #removed duplicate data from the dataframe.
print(df_1.head())
print(df_1.shape)


# In[8]:


df_1.isnull().sum()


# In[9]:


from openpyxl import load_workbook #Import load_workbook from openpyxl library so that after correcting the data can write it to other excel file.


# In[10]:


with pd.ExcelWriter('/home/deep/Desktop/KPMG_data.xlsx') as writer:  #first open the file in write mode.
    df_1.to_excel(writer, sheet_name='Transactions') #write a data to a sheet "Transactions" using to_excel function.
df_1_new = pd.read_excel("/home/deep/Desktop/KPMG_data.xlsx", 0)
df_1_new.head()


# In[11]:


df_2 = pd.read_excel("/home/deep/Downloads/KPMG_VI_New_raw_data_update_final.xlsx", 2)
df_2.head()


# In[12]:


print(df_2.shape) #Find out number of items we have
df_2.isnull().sum()


# In[13]:


col = ['Unnamed: 16','Unnamed: 17','Unnamed: 18','Unnamed: 19','Unnamed: 20']
df_2 = df_2.drop(col, axis=1)


# In[14]:


print(df_2.columns)
for i in df_2.columns:
    print(f"{df_2[i].value_counts()}")
    print("-" * 50)
print(df_2.describe())


# In[15]:


df_2['DOB'].sort_values(ascending=True)


# In[16]:


df_2.dropna(inplace=True)
print(df_2.shape)
df_2.drop_duplicates()
print(df_2.head())
print(df_2.shape)


# In[17]:


df_2.isnull().sum()


# In[18]:


with pd.ExcelWriter('/home/deep/Desktop/KPMG_data.xlsx', engine='openpyxl', mode='a') as writer:  
    df_2.to_excel(writer, sheet_name='NewCustomerList')
df_2_new = pd.read_excel("/home/deep/Desktop/KPMG_data.xlsx", 1)
df_2_new.head()


# In[19]:


df_3 = pd.read_excel("/home/deep/Downloads/KPMG_VI_New_raw_data_update_final.xlsx", 3)
df_3.head()


# In[20]:


print(df_3.columns)
for i in df_3.columns:
    print(f"{df_3[i].value_counts()}")
    print("-" * 50)
print(df_3.describe())


# In[21]:


df_3['DOB'].sort_values(ascending=True)


# In[22]:


print(df_3.shape) #Find out number of items we have
df_3.isnull().sum()


# In[23]:


df_3.dropna(inplace=True)
df_3.drop_duplicates()
print(df_3.head())
print(df_3.shape)


# In[24]:


df_3['gender']  = df_3['gender'].replace('M','Male').replace('F','Female').replace('Femal','Female')
df_3['gender'].value_counts()


# In[25]:


df_3.isnull().sum()


# In[26]:


del df_3['default'] #default column contains many garbage data.


# In[27]:


with pd.ExcelWriter('/home/deep/Desktop/KPMG_data.xlsx', engine='openpyxl', mode='a') as writer:  
    df_3.to_excel(writer, sheet_name='CustomerDemographic')
df_3_new = pd.read_excel("/home/deep/Desktop/KPMG_data.xlsx", 2)
df_3_new.head()


# In[28]:


for i in df_3_new.columns:
    print(f"{df_3_new[i].value_counts()}")
    print("-" * 50)


# In[29]:


df_3_new['DOB'].sort_values(ascending=True)


# In[30]:


df_3.columns


# In[31]:


df_4 = pd.read_excel("/home/deep/Downloads/KPMG_VI_New_raw_data_update_final.xlsx", 4)
print(df_4.head())
print(df_4.shape)


# In[32]:


df_4.isnull().sum()


# In[33]:


df_4.columns


# In[34]:


df_4.drop_duplicates()
print(df_4.shape)


# In[35]:


import datetime


# In[ ]:


curr_time_zone = datetime.datetime.now()
print(curr_time_zone)


# In[ ]:




