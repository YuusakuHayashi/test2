Attribute VB_Name = "RpaSystem"

Public Sub PrintOutSheet(ByRef v() As Variant)
    Dim printername As String: printername = v(0)
    Dim bookname As String: bookname = v(1)
    Dim sheetname As String: sheetname = v(2)
    Dim wbook As Workbook: Set wbook = Nothing
    Dim wsheet As Worksheet: Set wsheet = Nothing

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set wbook = Workbooks.Open(bookname, 0)
    wbook.Application.EnableEvents = False
    wbook.Application.DisplayAlerts = False
    wbook.Application.ScreenUpdating = False

    Set wsheet = wb.Worksheets(sheetname)

    If printername <> vbNullString Then
        wsheet.PrintOut ActivePrinter:=printername
    Else
        wsheet.PrintOut
    End If

    wbook.Application.ScreenUpdating = True
    wbook.Application.DisplayAlerts = True
    wbook.Application.EnableEvents = True
    wbook.Close

    Application.EnableEvents = True
    Application.DisplayAlerts = True

    Set wsheet = Nothing
    Set wbook = Nothing
End Sub

Public Function IsModuleExist(ByRef v() As Variant) As Variant
    Dim module As String: module = v(0)
    Dim ck As Boolean: ck = False
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        If vbc.Name = module Then
            ck = True
            Exit For
        End If
    Next
    If ck Then
        IsModuleExist = True
    Else
        IsModuleExist = False
    End If
End Function

Public Sub test()
    Dim v(1) As Variant
    Dim ans As Variant
    v(0) = "MacroImporter"
    ans = IsModuleExist(v)
End Sub