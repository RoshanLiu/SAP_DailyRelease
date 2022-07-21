import os
import Database




list = os.listdir(r'C://Users//roshan.liu//Scripts//DATA//SAP_OUTPUT')
path = 'C://Users//roshan.liu//Scripts//DATA//SAP_OUTPUT//'
for file in list:
    if file.split('-')[-1] == 'checked.txt':
        ele = file.split('-')
        newFile = ele[0]+'-'+ele[1]+'-'+ele[2]+'-solved.txt'
        os.rename(path+file,path+newFile)