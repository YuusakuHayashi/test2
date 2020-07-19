﻿Imports System.ComponentModel
Imports System.Threading

Public Class Window3
    Public Shared ServerName As System.String
    Public Shared DBName As System.String
    Public Shared SrcValue As System.String
    Public Shared DistValue As System.String
    Public Shared FieldName As System.String
    Public Shared TestDT As System.String
    Public Shared SqlVer As System.Int32
    Public Shared TableList() As System.Object

    Private excFlg As System.Boolean

    Private Sub Window3_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        test2.testmodule.defaultWinHeight = Me.Height
    End Sub


    Private Sub Window3_ContentRendered(sender As Object, e As EventArgs) Handles Me.ContentRendered
        Dim bd As New test2.testmodule.BData
        Dim i As System.Int32 : i = 0
        Dim res As System.String
        Dim dt As DateTime

        '-------------------------------------------
        bd.ServerName = Window3.ServerName
        bd.DBName = Window3.DBName
        bd.FieldName = Window3.FieldName
        bd.SrcValue = Window3.SrcValue
        bd.DistValue = Window3.DistValue
        bd.TestDT = Window3.TestDT
        bd.SqlVer = Window3.SqlVer
        bd.TableList = Window3.TableList
        '-------------------------------------------
        Try
            ThreadPool.QueueUserWorkItem(
                Sub()
                    Call Me.ControlsDisabled()
                    Me.excFlg = Constants.vbTrue

                    dt = DateTime.Now
                    Me.Result.Text = "START(" & dt.ToLocalTime & ")"

                    For j As System.Int32 = LBound(bd.TableList) To UBound(bd.TableList)
                        'TEST
                        'If Not tl(j) = "HYSTEST" Then
                        '    Continue For
                        'End If

                        bd.DTName = bd.TableList(j)("name")

                        Me.Dispatcher.BeginInvoke(
                            Sub()
                                Me.History.Content = "テーブル " & bd.DTName & " を処理しています・・・"
                            End Sub
                        )

                        res = test2.testmodule.Main2(bd)

                        Me.Dispatcher.BeginInvoke(
                            Sub()
                                Me.Progress.Content = j & "/" & UBound(bd.TableList) & "テーブルを処理しました"
                                Me.Result.AppendText(Constants.vbCrLf)
                                Me.Result.AppendText(res)
                                Me.Result.ScrollToEnd()
                            End Sub)
                    Next

                    Call test2.testmodule.TestTableDelete(bd)

                    Me.Dispatcher.BeginInvoke(
                        Sub()
                            dt = DateTime.Now
                            Me.Result.AppendText(Constants.vbCrLf)
                            Me.Result.AppendText("END(" & dt.ToLocalTime & ")")
                        End Sub
                    )

                    Call test2.testmodule.Save(bd)
                    bd = Nothing
                    Call Me.ControlsEnabled()
                    Me.excFlg = Constants.vbFalse
                    MsgBox("終了しました")

                End Sub, Nothing)

        Catch ex As Exception
            MsgBox("処理異常が発生しました")
        End Try

    End Sub

    Private Sub Back_Click(sender As Object, e As RoutedEventArgs) Handles Back.Click
        Call test2.testmodule.swichPages(Me, sender)
    End Sub

    Private Sub Window3_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Me.Title.FontSize = test2.testmodule.changeFontSize(20, e.NewSize.Height)
        Me.Contents.FontSize = test2.testmodule.changeFontSize(20, e.NewSize.Height)
        Me.History.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        Me.Progress.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
        Me.Result.FontSize = test2.testmodule.changeFontSize(10, e.NewSize.Height)
    End Sub

    Private Sub Window3_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        e.Cancel = Constants.vbFalse
        If Me.excFlg Then
            MsgBox("処理中は終了できません" & Constants.vbCrLf & "処理が終わるまでお待ちください")
            e.Cancel = Constants.vbTrue
        End If
    End Sub

    Private Sub ControlsEnabled()
        Me.Back.IsEnabled = Constants.vbFalse
    End Sub

    Private Sub ControlsDisabled()
        Me.Back.IsEnabled = Constants.vbTrue
    End Sub
End Class
