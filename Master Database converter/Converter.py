# -*- coding: utf-8 -*-
"""
Created on Fri Jun  4 14:51:38 2021

@author: z5035086
"""
import os
from csv import reader
import csv

from os import listdir
from os.path import isfile, join


OutputHeader = ["Survay Name", "Date Sampled", "Recorders", "Species","Data Type", "Plot ID", "Subplot", "Cover (%)", "Whatever A means", "Notes" ]

with open("output.csv", "w",newline='') as fp:
    wr = csv.writer(fp, dialect='excel')
    wr.writerow(OutputHeader)



FlieList = [f for f in listdir(os.getcwd()+'\\Input') if isfile(join(os.getcwd()+'\\Input', f))]


for F in FlieList:
    
    with open(os.getcwd()+'\\Input\\'+F, 'r') as read_obj:
        csv_reader = reader(read_obj)
        Data = list(csv_reader)
            
    Date = Data[0][1]
    SurvayName = Data[1][1]
    PlotNum = Data[2][1]
    Recorders = Data[3][1]
    DataType = Data[4][1]
    
    
    for r in Data[7:]:  # per row
        Species = r[0]
        Notes = r[1]     
           
        for c in [[2,3],[4,5],[6,7],[8,9]]: # pwe column
            NewRow = [SurvayName, Date, Recorders, Species, DataType, PlotNum,Data[6] [c[0]] [0],r [c[0]],r [c[1]], Notes]
        
            with open("output.csv", "a",newline='') as fp:
                wr = csv.writer(fp, dialect='excel')
                wr.writerow(NewRow)
                    
        
    
    
