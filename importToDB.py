import os
import Database




list = os.listdir(r'C://Users//roshan.liu//Scripts//DATA//SAP_OUTPUT')
path = 'C://Users//roshan.liu//Scripts//DATA//SAP_OUTPUT//'
for file in list:
    if file.split('.')[-1] == 'txt':
        Database.importFromText(path+file)