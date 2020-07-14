Imports System.Threading

Public Class Window3

    Public Shared ServerName As System.String
    Public Shared DBName As System.String
    Public Shared SrcValue As System.String
    Public Shared DistValue As System.String
    Public Shared FieldName As System.String


    Private Sub Window3_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        test2.testmodule.defaultWinHeight = Me.Height
    End Sub


    Private Sub Window3_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        Dim bd As New test2.testmodule.BData
        Dim tl As System.String()
        Dim i As System.Int32
        Dim res As System.String

        '-------------------------------------------
        bd.ServerName = Window3.ServerName
        bd.DBName = Window3.DBName
        bd.FieldName = Window3.FieldName
        bd.SrcValue = Window3.SrcValue
        bd.DistValue = Window3.DistValue
        '-------------------------------------------

        Me.Back.IsEnabled = False

        Me.result.Text = "START"
        ThreadPool.QueueUserWorkItem(
            Sub()
                tl = test2.testmodule.myTableList(bd)

                For i = LBound(tl) To UBound(tl)
                    ''--- TEST -------------------------
                    'If Not tl(i) = "HYSTEST" Then
                    '    Continue For
                    'End If
                    ''----------------------------------
                    bd.DTName = tl(i)
                    res = test2.testmodule.Main(bd)
                    Me.Dispatcher.BeginInvoke(
                        Sub()
                            Me.result.AppendText(Constants.vbCrLf)
                            Me.result.AppendText(res)
                            Me.result.ScrollToEnd()
                        End Sub)
                Next
                Me.Dispatcher.BeginInvoke(
                    Sub()
                        Me.result.AppendText(Constants.vbCrLf)
                        Me.result.AppendText("END")
                        MsgBox("終了しました")
                        Call test2.testmodule.Save(bd)
                        bd = Nothing
                        Me.Back.IsEnabled = True
                    End Sub
                )
            End Sub, Nothing)
    End Sub

    Private Sub Back_Click(sender As Object, e As RoutedEventArgs) Handles Back.Click
        Call test2.testmodule.swichPages(Me, sender)
    End Sub

    Private Sub Window3_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Me.Progress.FontSize = test2.testmodule.changeFontSize(36, e.NewSize.Height)
        Me.Contents.FontSize = test2.testmodule.changeFontSize(36, e.NewSize.Height)
        Me.history.FontSize = test2.testmodule.changeFontSize(12, e.NewSize.Height)
    End Sub
End Class
