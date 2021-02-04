Attribute VB_Name = "Rpa02"

Option Base 0

'README

Private Function GetInputDatas(ByRef v() As Variant) As String()
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim IIIDdata() As String

    Dim inf As String: inf = v(0)
    Dim inbook As Workbook
    Set inbook = Workbooks.Open(inf, 0)

    Dim insheet As Worksheet
    Set insheet = inbook.Activesheet

    Dim insheetnames() As String : insheetnames = v(1)
    Dim indatasstrings() As String : indatasstrings = v(2)
    Dim x As Integer : x = Ubound(insheetnames)
    Dim y As Integer : y = Ubound(indatasstrings)
    Dim z As Integer : z = Ubound(Split(indatasstrings(0), ","))
    ReDim IIIDdata(x, y, z)

    Dim x2 As Integer : x2 = 0
    Dim y2 As Integer : y2 = 0
    Dim z2 As Integer : z2 = 0
    For Each insheetname In insheetnames
        y2 = 0
        Set insheet = inbook.Worksheets(insheetname)
        For Each indatasstring In indatasstrings
            z2 = 0
            Dim indatas() As String : indatas = Split(indatasstring, ",")
            For Each indata In indatas
                IIIDdata(x2, y2, z2) = insheet.Range(indata).Text
                z2 = z2 + 1
            Next
            y2 = y2 + 1
        Next
        x2 = x2 + 1
    Next

    Set insheet = Nothing
    inbook.Close
    Set inbook = Nothing

    GetInputDatas = IIIDdata

    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Function

' HISTORY
'     2021-02-01: ゼロ値の空白置換対応
'     ＶＢ側で出力データを作成する際にゼロ値を空白置換するのは難しかったので、
'     エクセル側で行う 
Private Sub WriteOutputData(ByRef v() As Variant)
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim otf As String: otf = v(0)
    Dim obook As Workbook
    Set obook = Workbooks.Open(otf, 0)
    Dim osheet As Worksheet
    Set osheet = obook.ActiveSheet

    Dim indatasstrings() As String: indatasstrings = v(1)
    For Each indatasstring In indatasstrings
        Dim indatas() As String : indatas = Split(indatasstring, vbTab)
        Set osheet = obook.Worksheets(indatas(0))

        For i = 1 To Ubound(indatas) Step 2
            osheet.Range(indatas(i)).Value = indatas(i + 1)
            If osheet.Range(indatas(i)).Value = 0 Then
                osheet.Range(indatas(i)).Value = vbNullString
            End If
        Next 
    Next

    obook.Save
    Set osheet = Nothing
    obook.Close
    Set obook = Nothing

    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

