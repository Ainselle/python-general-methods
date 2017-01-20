
def writeExcel(listColumn,directory,column):
    print("Writing file...")
    import openpyxl
    try:
      xfile = openpyxl.load_workbook("./input/"+archivo)
      sheet = xfile.get_sheet_by_name('CasoPortonazo')
      for i in range(1,len(listaResp)+1):
          sheet[column+str(i+1)] = listaResp[i-1]
      xfile.save("./input/"+archivo)
    except:
      print("Writing failed")
    print("Writing completed.") 
      
