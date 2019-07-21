import  json, xlrd
def xlsParse(xlsfile):
    storage = []  
    datastore = []
    
    book = xlrd.open_workbook('NJ HMW Owners Rates 3.0 (Clean).xlsm')
    sh1 = book.sheet_by_index(0)
    sh1.cell_value(0,3)
    store = str(sh1.cell_value(8,1))
    ruleno, rulena = store.split('.',1)
    sh2 = book.sheet_by_index(7)
    for rx in range(1, sh2.nrows):
        
            datastore.append(sh2.row(rx)[0].value)
            data = {"Rule number": ruleno,
               "Rule name": rulena,
               "p1": sh2.row(rx)[2].value,
               "p3": sh2.row(rx)[3].value,
               
               }
            storage.append(data)

    
    out = json.dumps(storage, indent=4)
   
    f = open( 'xlsdata.json', 'w')
    f.write(out)
    
   
    



xlsParse('dataset.xls')

