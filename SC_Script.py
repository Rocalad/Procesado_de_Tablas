# -*- coding: utf-8 -*-
"""
Created on Wed Nov 10 08:04:13 2021

@author: rccalad
"""

import os
import pandas as pd
import re


# INPUTS

rootPath = r'G:\Mi unidad\Estudios Eléctricos\Proyectos\03. 3427-SP11 - Pozos Castilla\Rev. 0\02. Estudios\CC'  # Enter the main root where the whole folders are located in
singlePhaseFaultFolder = 'Monofásico'   # Enter the name of the folder where the results for single phase fault are located in
threePhaseFaultFolder = 'Trifásico'   # Enter the name of the folder where the results for three phase fault are located in
technicalParametersFileName = 'Paratec.txt'   # Enter the file's name where the tech parameters are contained followed by the .txt extension
resultsFileName = 'Resultados.xlsx' # Enter the final file's name followed by the .xlsx extension

# _________________________________________________________________________________________________________________________________________________________

# Function used to get the DataFrame merging the whole study cases for a specific type of fault

def GetDataFrameFromResults(rootPath, FaultFolderName):
    filesName=sorted(os.listdir(os.path.join(rootPath, FaultFolderName)))
    studyCases = {}
    
    for file in filesName:
        x=re.findall(r'C[0-9]+',file)
        caseName='cc_mono_' + x[0]
        caseName
        studyCases [caseName] = pd.read_csv(os.path.join(rootPath, FaultFolderName, file), sep='\t',header=0, usecols=('From','To','Ik"(L1)'), decimal = ',').rename(columns={'Ik"(L1)' : x[0], 'From': 'Barra'}).set_index('Barra')
        studyCases [caseName] = studyCases [caseName][studyCases [caseName]['To'].isnull()]
    
    table=pd.DataFrame()

    for name, dic in studyCases.items():
        table = table.merge(dic, how = 'outer', left_index=True , right_index = True).filter(regex='C[0-9]+').round(2)
    
    casesResultsKeys = ['C'+str(i) for i in range(len(table.keys()))]
    table=table[casesResultsKeys]
    
    return table

# Function used to get the final table for a specific type of fault

def GetFinalResultTable(GetDataFrameFromResults, rootPath, technicalParametersFileName, FaultFolderName):

    tableIndex=[index for index in GetDataFrameFromResults.index]
    techNodesParam = pd.read_csv(os.path.join(rootPath, technicalParametersFileName),sep='\t', usecols = ['Nombre', 'Vn', 'Zone'], skiprows = lambda x: x in [1, 1], decimal = ',' ).rename(columns={'Nombre' : 'Barra', 'Zone': 'Ubicación','Vn': 'Tensión nominal [kV]'}).set_index('Barra').loc[tableIndex].round(3)
    finalResults = techNodesParam.merge(GetDataFrameFromResults, how='right', left_index=True , right_index = True ).reset_index().sort_values(by=['Ubicación','Tensión nominal [kV]'],ascending=False).set_index(['Ubicación','Barra','Tensión nominal [kV]']).fillna('---')
    idx=[key for key in finalResults.keys()]
    finalResults.columns=pd.MultiIndex.from_product([['Ik"' + ' ' + FaultFolderName], idx])
    
    return finalResults

singlePhasefinalResults = GetFinalResultTable(GetDataFrameFromResults(rootPath, singlePhaseFaultFolder), rootPath, technicalParametersFileName, singlePhaseFaultFolder)

threePhasefinalResults = GetFinalResultTable(GetDataFrameFromResults(rootPath, threePhaseFaultFolder), rootPath, technicalParametersFileName, threePhaseFaultFolder)

# Getting combined results by merging both single and three phase results

combinedResults = singlePhasefinalResults.merge(threePhasefinalResults, how = 'outer', left_index=True, right_index = True)

# Exporting results to Excel

try:
    
    writer = pd.ExcelWriter(os.path.join(rootPath,resultsFileName), engine = 'xlsxwriter')
    
    singlePhasefinalResults.to_excel(writer, sheet_name = 'CC_1F')
    
    threePhasefinalResults.to_excel(writer, sheet_name = 'CC_3F')
    
    combinedResults.to_excel(writer, sheet_name = 'Resumen')
    
    writer.save()
    writer.close()
    
    print('The results file have been exported in the following location ---> ' + rootPath)
    
except: 
    
    print('Close the results file or rename it to create a new one.')

    
    







