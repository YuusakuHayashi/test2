Public Class Window3

    Dim test As test2.testclass
    Public Shared serverName As System.String
    Public Shared databaseName As System.String
    Public Shared sourceValue As System.String
    Public Shared distinationValue As System.String
    Public Shared fieldName As System.String


    Private Sub Window3_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        test = New testclass
        test.w = Me

        test2.testmodule.defaultWinHeight = Me.Height
    End Sub


    Private Async Sub Window3_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        Dim bd As New test2.testmodule.BData
        Dim tl As System.String()
        Dim i As System.Int32
        Dim res As System.String

        '-------------------------------------------
        bd.serverName = Window3.serverName
        bd.databaseName = Window3.databaseName
        bd.fieldName = Window3.fieldName
        bd.sourceValue = Window3.sourceValue
        bd.distinationValue = Window3.distinationValue
        '-------------------------------------------

        tl = test2.testmodule.TableList(bd)

        Me.result.Text = "START"

        For i = LBound(tl) To UBound(tl)
            ''--- TEST -------------------------
            'If Not tl(i) = "HYSTEST" Then
            '    Continue For
            'End If
            ''----------------------------------
            bd.datatableName = tl(i)
            res = test2.testmodule.Main(bd)
            Me.result.AppendText(Constants.vbCrLf)
            Me.result.AppendText(res)
            Me.history.Content = "実行中です・・・"
            Await Task.Delay(10)
        Next

        Me.result.AppendText(Constants.vbCrLf)
        Me.history.Content = "履歴"
        Me.result.AppendText("END")

        Call test2.testmodule.Save(bd)

        bd = Nothing
    End Sub

    Private Sub back_Click(sender As Object, e As RoutedEventArgs) Handles back.Click
        Call test2.testmodule.swichPages(test.w, sender)
    End Sub

    Private Sub Window3_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Me.history.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
    End Sub
End Class
