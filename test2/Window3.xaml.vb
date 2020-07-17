Imports System.ComponentModel
Imports System.Threading

Public Class Window3

    Public Shared ServerName As System.String
    Public Shared DBName As System.String
    Public Shared SrcValue As System.String
    Public Shared DistValue As System.String
    Public Shared FieldName As System.String
    Public Shared TestDT As System.String
    Public Shared SqlVer As System.Int32
    'Private endFlg As System.Boolean


    Private Sub Window3_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        test2.testmodule.defaultWinHeight = Me.Height
    End Sub


    Private Sub Window3_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        Dim bd As New test2.testmodule.BData
        Dim tl As System.String()
        Dim i As System.Int32 : i = 0
        Dim res As System.String
        Dim dt As DateTime

        'endFlg = 0

        '-------------------------------------------
        bd.ServerName = Window3.ServerName
        bd.DBName = Window3.DBName
        bd.FieldName = Window3.FieldName
        bd.SrcValue = Window3.SrcValue
        bd.DistValue = Window3.DistValue
        bd.TestDT = Window3.TestDT
        bd.SqlVer = Window3.SqlVer
        '-------------------------------------------

        Me.Back.IsEnabled = False

        dt = DateTime.Now
        Me.Result.Text = "START(" & dt.ToLocalTime & ")"

        ThreadPool.QueueUserWorkItem(
            Sub()
                tl = test2.testmodule.myTableList(bd)

                For j As System.Int32 = LBound(tl) To UBound(tl)
                    i += 1

                    'TEST
                    'If Not tl(j) = "HYSTEST" Then
                    '    Continue For
                    'End If

                    If tl(j) = bd.TestDT Then
                        Continue For
                    End If

                    bd.DTName = tl(j)

                    Me.Dispatcher.BeginInvoke(
                        Sub()
                            Me.History.Content = "テーブル " & bd.DTName & " を処理しています・・・"
                        End Sub
                    )

                    res = test2.testmodule.Main2(bd)

                    Me.Dispatcher.BeginInvoke(
                        Sub()
                            Me.Progress.Content = i & "/" & UBound(tl) & "テーブルを処理しました"
                            Me.Result.AppendText(Constants.vbCrLf)
                            Me.Result.AppendText(res)
                            Me.Result.ScrollToEnd()
                        End Sub)
                Next
                Me.Dispatcher.BeginInvoke(
                    Sub()
                        dt = DateTime.Now
                        Me.Result.AppendText(Constants.vbCrLf)
                        Me.Result.AppendText("END(" & dt.ToLocalTime & ")")
                        MsgBox("終了しました")
                        Call test2.testmodule.Save(bd)
                        bd = Nothing
                        Me.Back.IsEnabled = True
                    End Sub
                )
            End Sub, Nothing
        )

        'ThreadPool.QueueUserWorkItem(
        '    Sub()
        '        dt = DateTime.Now

        '        tl = test2.testmodule.myTableList(bd)

        '        Me.Dispatcher.BeginInvoke(
        '            Sub()
        '                Me.Back.IsEnabled = False
        '                Me.Result.Text = "START(" & dt.ToLocalTime & ")"
        '                For j As System.Int32 = LBound(tl) To UBound(tl)
        '                    i += 1

        '                    'TEST
        '                    'If Not tl(j) = "HYSTEST" Then
        '                    '    Continue For
        '                    'End If

        '                    If endFlg = 1 Then
        '                        endFlg = 2
        '                        Exit For
        '                    End If

        '                    If tl(j) = bd.TestDT Then
        '                        Continue For
        '                    End If

        '                    bd.DTName = tl(j)

        '                    Me.History.Content = "テーブル " & bd.DTName & " を処理しています・・・"

        '                    res = test2.testmodule.Main2(bd)

        '                    Me.Progress.Content = i & "/" & UBound(tl) & "テーブルを処理しました"
        '                    Me.Result.AppendText(Constants.vbCrLf)
        '                    Me.Result.AppendText(res)
        '                    Me.Result.ScrollToEnd()
        '                Next
        '                dt = DateTime.Now
        '                Me.Result.AppendText(Constants.vbCrLf)
        '                Me.Result.AppendText("END(" & dt.ToLocalTime & ")")
        '                Me.Back.IsEnabled = True
        '            End Sub)

        '        MsgBox("終了しました")

        '        Call test2.testmodule.Save(bd)

        '        bd = Nothing
        '        endFlg = 2
        '    End Sub, Nothing)
    End Sub

    Private Sub Back_Click(sender As Object, e As RoutedEventArgs) Handles Back.Click
        Call test2.testmodule.swichPages(Me, sender)
    End Sub

    Private Sub Window3_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Me.Progress.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        Me.Contents.FontSize = test2.testmodule.changeFontSize(36, e.NewSize.Height)
        Me.History.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        Me.Result.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
    End Sub

    'Private Sub Window3_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
    '    e.Cancel = Constants.vbTrue
    '    ThreadPool.QueueUserWorkItem(
    '        Sub()
    '            Me.Dispatcher.BeginInvoke(
    '                Sub()
    '                    Me.History.Content = "処理を終了しています。しばらくお待ちください・・・"
    '                    endFlg = 1
    '                    Do While True
    '                        Thread.Sleep(100)
    '                        If endFlg = 2 Then
    '                            Exit Do
    '                        End If
    '                    Loop
    '                    e.Cancel = Constants.vbFalse
    '                End Sub
    '            )
    '        End Sub, Nothing)
    'End Sub
End Class
