curDir = Replace(WScript.ScriptFullName, WScript.ScriptName, "") 
Set objExcel = CreateObject("Excel.Application")
Set w = objExcel.Workbooks.Open("C:\Users\Usuario\Desktop\DiptelPres 1.0\salida\INE-1600.xlsx")
objExcel.Application.DisplayAlerts = False
objExcel.Application.Visible = True
Set h = w.ActiveSheet
h.PageSetup.CenterHeaderPicture.Filename = curDir + "encabezado.png"
h.PageSetup.CenterHeaderPicture.Width = 550
h.PageSetup.CenterHeader = "&G"
h.PageSetup.CenterFooterPicture.Filename = curDir + "pie.png"
h.PageSetup.CenterFooterPicture.Width = 550
h.PageSetup.CenterFooter = "&G"
h.cells(488,2).Select
call h.Pictures.Insert(curDir + "firma.png")
w.Save
Call h.ExportAsFixedFormat(xlTypePDF, "C:\Users\Usuario\Desktop\DiptelPres 1.0\salida\INE-1600.pdf" , xlQualityMinimum, True, False, , , True)
w.Close