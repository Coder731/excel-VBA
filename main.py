import win32com.client
import pandas as pd

excel = win32com.client.Dispatch("Excel.Application")
workbook = excel.Workbooks.Open("{}{}.xlsm".format(path, file), True, True)

dict_modules = {}
for i in workbook.VBProject.VBComponents:
    name = i.name
    lines = workbook.VBProject.VBComponents(name).CodeModule.CountOfLines

    # To jump empty modules
    if lines == 0:
        pass
    else:
        text = workbook.VBProject.VBComponents(name).CodeModule.Lines(1,lines)
        dict_modules[name] = [text]

df = pd.DataFrame(dict_modules)
