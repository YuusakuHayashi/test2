Attribute VB_Name = "RpaLibrary"
Private Const FSOBJ = "Scripting.FileSystemObject"

Type HankakuZenkaku
    rngs As Object          '--- Range Collection
    rng As Range            '--- 各々のセル範囲
    r As Range
    conversion As Integer
    oxk_mode As Integer
End Type


Private Sub set_group_underline(ByRef BGN_CEL As Range)
    Dim cel_val As Variant
    Dim lastrow As Long: lastrow = Cells(ActiveSheet.Rows.Count, BGN_CEL.Column).End(xlUp).Row
    Dim lastcol As Long: lastcol = Cells(BGN_CEL.Row, ActiveSheet.Columns.Count).End(xlToLeft).Column

    cel_val = BGN_CEL.Value
    Dim i As Long: i = BGN_CEL.Row
    Dim j As Long: j = BGN_CEL.Column

    Do
        If Cells(i, j).Value <> cel_val Then
            Range(Cells(i, 1), Cells(i, lastcol)).Borders(xlEdgeTop).LineStyle = xlContinuous
            cel_val = Cells(i, j).Value
        End If
        i = i + 1
    Loop Until Cells(i, j).Row = lastrow
End Sub

Private Sub sample()
    ' This is a sample code
    Call set_group_underline(Cells(2, 3))
End Sub

Private Sub set_highlights(ByRef BGN_CEL As Range, ByVal TGT_VAL As Variant, ByVal c_idx)
    ' about colorindex
    '   https://www.sejuku.net/blog/32288

    Dim cel_val As Variant
    Dim lastrow As Long: lastrow = Cells(ActiveSheet.Rows.Count, BGN_CEL.Column).End(xlUp).Row
    Dim lastcol As Long: lastcol = Cells(BGN_CEL.Row, ActiveSheet.Columns.Count).End(xlToLeft).Column

    cel_val = TGT_VAL
    Dim i As Long: i = BGN_CEL.Row
    Dim j As Long: j = BGN_CEL.Column

    Do
        If Cells(i, j).Value = cel_val Then
            Range(Cells(i, 1), Cells(i, lastcol)).Interior.ColorIndex = i_idx
        End If
        i = i + 1
    Loop Until Cells(i, j).Row = lastrow
End Sub

Public Function GetStringArrayOfFile(ByVal fn As String) As String()
    Dim fso As Object: Set fso = Nothing
    Dim ts As Object: Set ts = Nothing
    Dim i As Integer: i = 0
    Dim tmp() As String

    Set fso = CreateObject(FSOBJ)
    Set ts = fso.OpenTextFile(fn, 1, True, -2)

    ReDim tmp(100)
    Do Until ts.AtEndOfStream
        tmp(i) = ts.ReadLine
        i = i + 1
    Loop
    ReDim Preserve tmp(i - 1)

    GetStringArrayOfFile = tmp
    Set ts = Nothing
    Set fso = Nothing
End Function

Public Function GetStringArrayOfFile2(ByVal fn As String) As String()
    Dim fso As Object: Set fso = Nothing
    Dim ts As Object: Set ts = Nothing
    Dim i As Integer: i = 0
    Dim tmp() As String

    Set fso = CreateObject(FSOBJ)
    Set ts = fso.OpenTextFile(fn, 1, True, -2)

    tmp = Split(ts.ReadAll(), vbCrlf)

    GetStringArrayOfFile2 = tmp
    Set ts = Nothing
    Set fso = Nothing
End Function

'--------------------------------------------------------------------------------------------------

Public Sub ConvertZenkakuToHankaku()
    Call ConvertHankakuZenkaku(vbNarrow)
End Sub

Public Sub ConvertHankakuToZenkaku()
    Call ConvertHankakuZenkaku(vbWide)
End Sub

Public Sub ConvertKanjiToHiragana()
    Call ConvertHankakuZenkaku(20)
End Sub

Public Sub ConvertKanjiToHiraganaToHankaku()
    Call ConvertHankakuZenkaku(20, vbNarrow)
End Sub

Public Function ConvertToOmoji(ByVal txt As String) As String
    txt = Replace(txt, "ァ", "ア")
    txt = Replace(txt, "ィ", "イ")
    txt = Replace(txt, "ゥ", "ウ")
    txt = Replace(txt, "ェ", "エ")
    txt = Replace(txt, "ォ", "オ")
    txt = Replace(txt, "ャ", "ヤ")
    txt = Replace(txt, "ュ", "ユ")
    txt = Replace(txt, "ョ", "ヨ")
    txt = Replace(txt, "ッ", "ツ")
    txt = Replace(txt, "ｧ", "ｱ")
    txt = Replace(txt, "ｨ", "ｲ")
    txt = Replace(txt, "ｩ", "ｳ")
    txt = Replace(txt, "ｪ", "ｴ")
    txt = Replace(txt, "ｫ", "ｵ")
    txt = Replace(txt, "ｬ", "ﾔ")
    txt = Replace(txt, "ｭ", "ﾕ")
    txt = Replace(txt, "ｮ", "ﾖ")
    txt = Replace(txt, "ｯ", "ﾂ")
    ConvertToOmoji = txt
End Function

Private Sub ConvertHankakuZenkaku(ParamArray conversions() As Variant)
    Dim hz As HankakuZenkaku
    Dim txt As String
    Dim i As Integer
    
    Set hz.rngs = Selection.Areas

    For i = LBound(conversions) To UBound(conversions)
        hz.conversion = conversions(i)
        For Each hz.rng In hz.rngs
            hz = CheckValidRange(hz)
            If Not hz.rng Is Nothing Then
                For Each hz.r In hz.rng
                    If hz.r.Text <> vbNullString Then
                        txt = hz.r.Text
                        Select Case hz.conversion
                            Case vbUpperCase, _
                                    vbLowerCase, _
                                    vbProperCase, _
                                    vbWide, _
                                    vbNarrow, _
                                    vbKatakana, _
                                    vbHiragana, _
                                    vbUnicode, _
                                    vbFromUnicode
                                txt = StrConv(txt, hz.conversion)
                            Case 10
                                txt = ConvertToOmoji(txt)
                            Case 20
                                txt = Application.GetPhonetic(txt)
                            Case Else
                                txt = txt
                        End Select
                        hz.r.Value = txt
                    End If
                Next
                Set hz.r = Nothing
            End If
            Set hz.rng = Nothing
        Next
    Next
    Set hz.rngs = Nothing
End Sub

Private Function CheckValidRange(ByRef hz As HankakuZenkaku) As HankakuZenkaku
    Dim min_row As Integer
    Dim min_col As Integer
    Dim max_row As Long
    Dim max_col As Long

    If hz.rng.Address = hz.rng.Worksheet.Cells.Address Then
        hz.rng = Nothing
    Else
        min_row = hz.rng.Row
        min_col = hz.rng.Column
        max_row = hz.rng.Row + hz.rng.Rows.Count - 1
        max_col = hz.rng.Column + hz.rng.Columns.Count - 1
    
        If hz.rng.Address = hz.rng.EntireColumn.Address Then
            max_row = 1
            Dim x As Integer
            For x = 1 To hz.rng.Columns.Count
                If max_row < hz.rng.Cells(hz.rng.Rows.Count, x).End(xlUp).Row Then
                    max_row = hz.rng.Cells(hz.rng.Rows.Count, x).End(xlUp).Row
                End If
            Next
        End If
    
        If hz.rng.Address = hz.rng.EntireRow.Address Then
            max_col = 1
            Dim y As Integer
            For y = 1 To hz.rng.Rows.Count
                If max_col < hz.rng.Cells(y, hz.rng.Columns.Count).End(xlToLeft).Column Then
                    max_col = hz.rng.Cells(y, hz.rng.Columns.Count).End(xlToLeft).Column
                End If
            Next
        End If
        Set hz.rng = hz.rng.Worksheet.Range( _
                          hz.rng.Worksheet.Cells(min_row, min_col), _
                          hz.rng.Worksheet.Cells(max_row, max_col))
    End If
    CheckValidRange = hz
End Function
'--------------------------------------------------------------------------------------------------
