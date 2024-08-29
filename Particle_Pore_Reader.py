"""
Created on Sun Feb 27 22:05:12 2022

@author: sivakumarvalluri
"""
# ImageJ csv results file particle cross-section data reader

import pandas as pd
import os
import glob
import numpy as np  
  
# Renaming xls files as csv files #Misreads particle size when imageJ files from xls converetd to csv
# change location as needed and be mindful of slashes
#folder = 'C:/Users/sivak/Desktop/Outcome'
#for filename in glob.iglob(os.path.join(folder, '*.xls')):
    #os.rename(filename, filename[:-4] + '.csv')
    #print('renamed')

# Change location as needed and dont forget double back slashes
os.chdir('D:\\Postdoctoral Work\\Experimental Data\\Project III- Study of porous Al-MoO3-KNO3\\Phi-3 (2.83)\\SEM images\\Particle Cross-Section\\ImageJ Analysis of Microstructure\\Output')
path = os.getcwd()
csv_files = glob.glob(os.path.join(path, "*.csv"))
  
# Master database initializing
Particle_df = pd.DataFrame() 

  
# loop over the list of csv files in folder
for f in csv_files:
    
    # Reading the csv file
    df = pd.read_csv(f)
    #df = pd.read_csv(f, sep='\s+', engine='python') #Misreads particle size when imageJ files from xls converetd to csv
      
    # Details of a particle of given size 
    print('File being processed:', f.split("\\")[-1])
    number_of_pores = len(df.index)-1
    particle_size=(df.loc[number_of_pores,'Major']+df.loc[number_of_pores,'Minor'])/2
    porosity= (df.loc[ 0:len(df.index)-2,["Area"]].sum())/(df.loc[number_of_pores,'Area'])
    pore_sizes=(df['Major']+df['Minor'])/2
    pore_sizes=pore_sizes.drop(index=len(pore_sizes.index)-1)
    
    # Reading data into master database
    new_row=pd.Series([particle_size,number_of_pores,porosity.loc['Area'],pore_sizes])
    row_df=pd.DataFrame([new_row])
    Particle_df=pd.concat([row_df,Particle_df],ignore_index=True)
    print('done')
  # out of loop  
    

# Master database headers
Particle_df.columns =['Particle_Size', 'Pore_number', 'Porosity', 'Pore_Size']



# New dataframe for particle-pore size relationship
n=np.arange(-10,25,1)
size=1.3**n

Data_df = pd.DataFrame() 
Data_df.insert(0, "Size_bin", size, True)
Data_df["Particle_Count"]=""
Data_df["Pore_number"]=""
Data_df["Porosity"]=""
Data_df["Pore_Size"]=""

print('Binning detail based on particle size')

for j in range(0,len(Data_df["Size_bin"])-2,1):
    AA=[i for i in range(len(Particle_df["Particle_Size"])) if ((Particle_df.Particle_Size[i] >= Data_df.Size_bin[j]) and (Particle_df.Particle_Size[i] < Data_df.Size_bin[j+1]))]
    for k in range(0,len(AA)):
        
        #Number of particles in bin
        Data_df.at[j,'Particle_Count']=len(AA)

        #Pore sizes 
        A1=pd.Series(Particle_df.at[AA[k],'Pore_Size'])
        A2=pd.Series(Data_df.at[j,'Pore_Size'])
        A12=pd.concat([A1, A2], ignore_index=True)
        Data_df.at[j,'Pore_Size']=A12
        

        #Porosity
        B1=pd.Series(Particle_df.at[AA[k],'Porosity'])
        B2=pd.Series(Data_df.at[j,'Porosity'])
        B12=pd.concat([B1, B2], ignore_index=True)
        Data_df.at[j,'Porosity']=B12
        
 
        #Pore number
        C1=pd.Series(Particle_df.at[AA[k],'Pore_number'])
        C2=pd.Series(Data_df.at[j,'Pore_number'])
        C12=pd.concat([C1, C2], ignore_index=True)
        Data_df.at[j,'Pore_number']=C12
        
print('Binning complete')

print('Writing onto excel file...')

#Representative size (average of successive bins)
title_size=[0]*len(Data_df['Size_bin'])
bins=Data_df["Size_bin"].to_list()
for i in range(0,len(Data_df["Size_bin"])-2,1):
         title_size[i+1]=(bins[i]+bins[i+1])/2

#change iteration range if needed

import xlsxwriter
with xlsxwriter.Workbook('Pore number.xlsx') as workbook:  
    worksheet = workbook.add_worksheet("Pore Number")
    for i in range(0,len(title_size),1):    
             worksheet.write(1,i, title_size[i])
    for m in range(0,len(Data_df["Size_bin"])-1,1):
             DAM = Data_df.at[m,'Pore_number']
             for n in range(0,len(DAM)-1,1):
                      worksheet.write(n+2,m, DAM[n])
    

with xlsxwriter.Workbook('Pore Size.xlsx') as workbook2:  
    worksheet2 = workbook2.add_worksheet("Pore Size")
    for i in range(0,len(title_size),1):    
             worksheet2.write(1,i, title_size[i])
    for m in range(0,len(Data_df["Size_bin"])-1,1):
             DAM = Data_df.at[m,'Pore_Size']
             for n in range(0,len(DAM)-1,1):
                      worksheet2.write(n+2,m, DAM[n])         

with xlsxwriter.Workbook('Porosity.xlsx') as workbook3:  
    worksheet3 = workbook3.add_worksheet("Porosity")
    for i in range(0,len(title_size),1):    
             worksheet3.write(1,i, title_size[i])
    for m in range(0,len(Data_df["Size_bin"])-1,1):
             DAM = Data_df.at[m,'Porosity']
             for n in range(0,len(DAM)-1,1):
                      worksheet3.write(n+2,m, DAM[n])        
                      
print('Let the games begin')                      