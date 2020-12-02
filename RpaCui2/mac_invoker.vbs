Dim exapp
Dim exbok
Dim fso
Dim pth
Set exapp = WScript.CreateObject("Excel.Application")
Set fso   = WScript.CreateObject("Scripting.FileSystemObject")
pth = fso.GetFolder(".").Path
Set exbok = exapp.Workbooks.Open(WScript.Arguments(0), 0)
exapp.Visible = False
exapp.Run WScript.Arguments(1) + "." + WScript.Arguments(2)
exbok.Close()
exapp.quit
Set fso   = Nothing
Set exbok = Nothing
Set exapp = Nothing
