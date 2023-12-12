import win32com.client
xlApp = win32com.client.Dispatch("Excel.Application")
xlApp.visible = True

work_file = r'\\Lxjjaq6628\d\prg_sandbox\Eng_salesForecast\aaa.xlsx'
list_wb = xlApp.Workbooks.Open(work_file)

sh = list_wb.Worksheets("sheet1")
sh.Activate()