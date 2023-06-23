#!/usr/bin/env python
# coding: utf-8

# Requirements to run this code file :-
# 
# 1. Must have a ChromeDriver installed in the C drive of your PC.
# 2. Make a new folder in your desktop as Trove which can be used as a dowload directory further.

# In[1]:


get_ipython().system('pip install selenium')


# In[2]:


from selenium import webdriver


# In[3]:


from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# In[4]:


import os
import time
import pandas as pd


# In[5]:


# Set the download directory path
download_directory = (r"C:\Users\Subho_98\Desktop\Trove")

# Set Chrome options
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_experimental_option('prefs', {
    'download.default_directory': download_directory,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': True,
    'plugins.always_open_pdf_externally': True, # always open pdf externally
    'download.open_pdf_in_system_reader': False, # don't open pdf in system reader
    'profile.default_content_settings.popups': 0, # disable popups
})


# In[6]:


# Pass the Chrome options to the WebDriver
driver = webdriver.Chrome('C:\Chrome', options=chrome_options)
driver.get('https://biocarbonregistry.com/en/projects/')


# In[7]:


button = driver.find_element(By.XPATH, '//*[@id="tabla-iniciativas_wrapper"]/div[1]/a[1]')
button.click()
print("Project list downloaded!!!")


# In[8]:


files = r"C:\Users\Subho_98\Desktop\Trove\Projects – BioCarbon Registry.xlsx"


# In[9]:


df = pd.read_excel(files, header=1)
df


# In[10]:


base_url = "https://app.biocarbonregistry.com/summary-report/"
download_path = (r"C:\Users\Subho_98\Desktop\Trove")


# In[11]:


x=0
while x<len(df):
    value = df.iloc[x,0]
    folder_path = os.path.join(download_path,str(value))
    os.makedirs(folder_path, exist_ok=True)
                               
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")  # Run Chrome in headless mode
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_experimental_option('prefs', {
        'download.default_directory': folder_path,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True,
        'plugins.always_open_pdf_externally': True, # always open pdf externally
        'download.open_pdf_in_system_reader': False, # don't open pdf in system reader
        'profile.default_content_settings.popups': 0, # disable popups
    })

    driver = webdriver.Chrome('C:\Chrome', options=chrome_options)
    driver.get(base_url + str(value) + '/en')
    time.sleep(3)

    x = x+1
    if x>len(df):
        print("Downloaded!!")
    


# In[13]:


x=0
file_name_common = ("Información general del proyecto.pdf")
while x<len(df):
    value = df.iloc[x,0]
    old_name = (os.path.join(download_path, str(value), file_name_common))
    new_name = (os.path.join(download_path, str(value), "Bio_Carbon_" +str(value)+ "_" + file_name_common))
    os.rename(old_name, new_name)
    x = x+1


# In[14]:


Old_Name = ('_Información general del proyecto.pdf')
Document_Type = ('Summary report')
Registry = 'Bio_Carbon_'
df['Old_Name'] = Old_Name
df['Document_Type'] = Document_Type
df['New_Name'] = Registry + df['#' ].astype(str) + Old_Name


# In[15]:


df


# In[16]:


#Saving the updated registry
df.to_excel(download_path + r'\BioCarbonUpdated.xlsx', index=False)


# In[ ]:




