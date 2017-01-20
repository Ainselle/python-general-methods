
def writeExcel(listColumn,directory,column):
    print("Writing file...")
    import openpyxl
    try:
      xfile = openpyxl.load_workbook("./input/"+archivo)
      sheet = xfile.get_sheet_by_name('CasoPortonazo')
      for i in range(1,len(listColumn)+1):              #starting from second row(avoid overwriting the headers)
          sheet[column+str(i+1)] = listColumn[i-1]
      xfile.save("./input/"+archivo)
    except:
      print("Writing failed")
    print("Writing completed.") 
      
