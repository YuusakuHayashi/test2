Imports System.IO

Public Class ExitCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

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

        Dim jh As New RpaCui.JsonHandler(Of List(Of RpaUpdater))
        Dim rrus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)
        Dim uru As RpaUpdater
        rrus.Sort(
            Function(before, after)
                Return (before.ReleaseDate < after.ReleaseDate)
            End Function
        )
        uru = rrus.Last

        ' リテラル部分は後で修正する（面倒なので、リテラルで仮置きしている）
        '-----------------------------------------------------------------------------------------'
        dat.Transaction.AutoCommandText.Add($"updaterobot {uru.ReleaseId}")
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

        'If dat.Project IsNot Nothing Then
        '    If File.Exists(dat.Project.SystemJsonChangedFileName) Then
        '        Dim yorn As String = vbNullString
        '        Do
        '            yorn = vbNullString
        '            Console.WriteLine()
        '            Console.WriteLine($"Projectに変更があります。変更を保存しますか？ (y/n)")
        '            yorn = dat.Transaction.ShowRpaIndicator(dat)
        '        Loop Until yorn = "y" Or yorn = "n"
        '        If yorn = "y" Then
        '            RpaModule.Save(dat.Project.SystemJsonFileName, dat.Project, dat.Project.SystemJsonChangedFileName)
        '        Else
        '            File.Delete(dat.Project.SystemJsonChangedFileName)
        '        End If
        '    End If
        'End If

        'If File.Exists(RpaInitializer.SystemIniChangedFileName) Then
        '    Dim yorn2 As String = vbNullString
        '    Do
        '        yorn2 = vbNullString
        '        Console.WriteLine()
        '        Console.WriteLine($"Initializerに変更があります。変更を保存しますか？ (y/n)")
        '        yorn2 = dat.Transaction.ShowRpaIndicator(dat)
        '    Loop Until yorn2 = "y" Or yorn2 = "n"
        '    If yorn2 = "y" Then
        '        RpaModule.Save(RpaCui.SystemIniFileName, dat.Initializer, RpaInitializer.SystemIniChangedFileName)
        '    Else
        '        File.Delete(RpaInitializer.SystemIniChangedFileName)
        '    End If
        'End If

        dat.Transaction.ExitFlag = True
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
