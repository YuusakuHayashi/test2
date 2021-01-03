Imports System.IO
Imports Rpa00

Public Class RpaPrinterUtility : Inherits RpaUtilityBase

    Public PrintText As String
    Public PrintFile As FileInfo

    Public Overrides ReadOnly Property ExecuteHandler(dat As Object) As Object
        Get
            Return Nothing
        End Get
    End Property

    Public Overloads Sub TextPrintRequest(ByVal txt As String)
        Me.PrintText = txt
        Dim pd As New Drawing.Printing.PrintDocument
        AddHandler pd.PrintPage, AddressOf _TextPrintOut

        pd.Print()
    End Sub

    Public Overloads Sub TextPrintRequest(ByRef fi As FileInfo, ByVal fenc As String)
        Dim sr As StreamReader
        sr = New StreamReader(fi.FullName, System.Text.Encoding.GetEncoding(fenc))

        Call TextPrintRequest(sr.ReadToEnd)

        If sr IsNot Nothing Then
            sr.Close()
            sr.Dispose()
        End If
    End Sub

    Private Sub _TextPrintOut(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        Dim pos = 0
        Dim txt = Me.PrintText
        Dim y = e.MarginBounds.Top
        Dim x = e.MarginBounds.Left
        Dim font = New Drawing.Font("ＭＳ ゴシック", 10)
        Dim line = vbNullString

        txt = txt.Replace(vbCrLf, vbLf)
        txt = txt.Replace(vbCr, vbLf)

        While (e.MarginBounds.Bottom > y + font.Height) And (pos < txt.Length)
            line = vbNullString
            While True
                If (pos >= txt.Length) Or txt.Chars(pos) = vbLf Then
                    pos += 1
                    Exit While
                End If

                line += txt.Chars(pos)
                If e.Graphics.MeasureString(line, font).Width > e.MarginBounds.Width Then
                    line = line.Substring(0, line.Length - 1)
                    Exit While
                End If

                pos += 1
            End While
            e.Graphics.DrawString(line, font, Drawing.Brushes.Black, x, y)
            y += font.GetHeight(e.Graphics)
        End While

        If pos >= txt.Length Then
            e.HasMorePages = False
            pos = 0
        Else
            e.HasMorePages = True
        End If
    End Sub
End Class
