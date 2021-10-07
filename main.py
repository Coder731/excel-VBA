## PART 1:

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

## PART 2:

# To get the full text
module_name = df["module_name"][0]
#To get by line
module_text_by_line = module_name.splitlines()
