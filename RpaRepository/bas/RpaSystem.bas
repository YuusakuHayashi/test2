Attribute VB_Name = "RpaSystem"
Option Explicit

Private Function IsSheetExist(ByRef v() As Variant) As Boolean
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim b As Boolean : b = False

    Dim bookname As String: bookname = v(0)
    Dim sheetname As String: sheetname = v(1)
    Dim wbook As Workbook: Set wbook = Nothing
    Set wbook = Workbooks.Open(bookname, 0)

    Dim ws As Worksheet: Set ws = Nothing
    For Each ws In wbook.Worksheets
        If ws.Name = sheetname Then
            b = True
            Exit For
        End If
    Next

    wbook.Close
    Set wbook = Nothing

    IsSheetExist = b

    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Function

Public Sub PrintOutSheet(ByRef v() As Variant)
    Dim printername As String: printername = v(0)
    Dim bookname As String: bookname = v(1)
    Dim sheetname As String: sheetname = v(2)
    Dim wbook As Workbook: Set wbook = Nothing
    Dim wsheet As Worksheet: Set wsheet = Nothing

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set wbook = Workbooks.Open(bookname, 0)
    'wbook.Application.EnableEvents = False
    'wbook.Application.DisplayAlerts = False
    'wbook.Application.ScreenUpdating = False

    Set wsheet = wbook.Worksheets(sheetname)

    If printername <> vbNullString Then
        wsheet.PrintOut ActivePrinter:=printername
    Else
        wsheet.PrintOut
    End If

    'wbook.Application.ScreenUpdating = True
    'wbook.Application.DisplayAlerts = True
    'wbook.Application.EnableEvents = True
    wbook.Close

    Application.EnableEvents = True
    Application.DisplayAlerts = True

    Set wsheet = Nothing
    Set wbook = Nothing
End Sub

Private Function IsModuleExist(ByRef v() As Variant) As Boolean
    Dim module As String: module = v(0)
    Dim ck As Boolean: ck = False
    Dim vbc As Object

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

    Set vbc = Nothing
End Function

Private Function GetCellText(ByRef v() As Variant) As String
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim inf As String: inf = v(0)
    Dim inbook As Workbook
    Set inbook = Workbooks.Open(inf, 0)

    Dim ins As String: ins = v(1)
    Dim insheet As Worksheet
    Set insheet = inbook.Worksheets(ins)

    Dim inc As String : inc = v(2)
    Dim incell As Range
    Set incell = insheet.Range(inc)

    GetCellText = incell.Text

    Set incell = Nothing
    Set insheet = Nothing
    inbook.Close
    Set inbook = Nothing

    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Function

Public Sub test()
    Dim v(1) As Variant
    Dim ans As Variant
    v(0) = "MacroImporter"
    Call IsModuleExist(v)
End Sub
