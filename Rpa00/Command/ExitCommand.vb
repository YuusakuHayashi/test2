Imports System.IO

Public Class ExitCommand : Inherits RpaCommandBase
    '' IntranetClientServerProjectで実行
    Private Function UpdateCheck(ByRef dat As RpaDataWrapper) As Integer
        If Not dat.Project.IsUpdateAvailable Then
            Return 0
        End If

        Dim idx As String = vbNullString
        If dat.Project.AutoUpdate Then
            idx = "0"
        Else
            Do
                idx = vbNullString
                Console.WriteLine()
                Console.WriteLine($"0 ...更新して終了")
                Console.WriteLine($"1 ...終了")
                idx = dat.Transaction.ShowRpaIndicator(dat)
                If idx = "0" Or idx = "1" Then
                    Exit Do
                End If
            Loop Until False
        End If

        If idx = "1" Then
            Return 0
        End If

        Dim jh As New Rpa00.JsonHandler(Of List(Of RpaUpdater))
        Dim rrus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)
        Dim uru As RpaUpdater
        rrus.Sort(
            Function(before, after)
                If after.IsCritical Then
                    Return before.ReleaseDate < after.ReleaseDate
                End If
                Return False
            End Function
        )
        uru = rrus.Last

        ' リテラル部分は後で修正する（面倒なので、リテラルで仮置きしている）
        '-----------------------------------------------------------------------------------------'
        dat.System.LateBindingCommands.Add($"updateproject {uru.ReleaseId}")
        '-----------------------------------------------------------------------------------------'

        Return 0
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        ' アップデートチェック
        '-----------------------------------------------------------------------------------------'
        Dim i As Integer
        If dat.Project.SystemArchTypeName = (New IntranetClientServerProject).GetType.Name Then
            i = UpdateCheck(dat)
        End If
        ''-----------------------------------------------------------------------------------------'

        dat.System.ExitFlag = True
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecuteIfNoProject = True
    End Sub
End Class
