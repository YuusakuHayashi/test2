Imports Rpa00
Imports System.IO

' IntranetClientServerProject専用コマンド

Public Class AttachRobotCommand : Inherits RpaCommandBase
    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        Dim rpa = dat.Project.DeepCopy
        Dim err1 As String = $"'MyDirectory'が設定されていません"
        If String.IsNullOrEmpty(rpa.MyDirectory) Then
            Console.WriteLine(err1)
            Console.WriteLine()
            Return False
        End If

        Dim err2 As String = $"MyDirectory '{rpa.MyDirectory}' が存在しません"
        If Not Directory.Exists(rpa.MyDirectory) Then
            Console.WriteLine(err2)
            Console.WriteLine()
            Return False
        End If

        If dat.Transaction.Parameters.Count = 0 Then
            Dim err3 As String = $"プロジェクトにロボットが存在しません"
            If dat.Project.RobotAliasDictionary.Count = 0 Then
                Console.WriteLine(err3)
                Console.WriteLine()
                Return False
            End If
        End If

        If dat.Transaction.Parameters.Count > 0 Then
            rpa.RobotName = dat.Transaction.Parameters(0)

            Dim err4 As String = $"'RootRobotDirectory'が設定されていません"
            If String.IsNullOrEmpty(rpa.RootRobotDirectory) Then
                Console.WriteLine(err4)
                Console.WriteLine()
                Return False
            End If

            Dim err5 As String = $"RootRobotDirectory '{rpa.RootRobotDirectory}' が存在しません"
            If Not Directory.Exists(rpa.RootRobotDirectory) Then
                Console.WriteLine(err5)
                Console.WriteLine()
                Return False
            End If

            Dim err6 As String = $"'{rpa.RootRobotDirectory}' に '{rpa.RootRobotIniFileName}' がありません"
            If Not File.Exists(rpa.RootRobotIniFileName) Then
                Console.WriteLine(err6)
                Console.WriteLine()
                Return False
            End If
        End If

        Return True
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim roboname As String = vbNullString

        ' Parameters(0) = アタッチするロボットの指定の有無
        If dat.Transaction.Parameters.Count > 0 Then
            roboname = dat.Transaction.Parameters(0)
        Else
            Dim idx As Integer = -1
            Dim lastidx As Integer = -1
            Dim dic As Dictionary(Of String, String) = dat.Project.RobotAliasDictionary
            Dim pairs As List(Of KeyValuePair(Of String, String)) = dic.ToList
            For Each robo In pairs
                idx = pairs.IndexOf(robo)
                Console.WriteLine($"{idx}   ロボット名:{robo.Key}   登録名:{robo.Value}")
            Next
            lastidx = idx + 1
            Console.WriteLine($"{lastidx}   やっぱりやめる")
            Console.WriteLine()

            Dim idx2 As Integer = -1
            Dim idxtext As String = vbNullString
            Do
                Console.WriteLine($"アタッチするロボットを選択してください")
                idxtext = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If IsNumeric(idxtext) Then
                    idx2 = Integer.Parse(idxtext)
                Else
                    idx2 = lastidx + 1
                End If
            Loop Until idx2 <= lastidx

            If idx2 = lastidx Then
                Return 0
            Else
                roboname = pairs(idx2).Key
            End If
        End If

        If Not dat.Project.RobotAliasDictionary.ContainsKey(roboname) Then
            dat.Project.RobotAliasDictionary.Add(roboname, roboname)
        End If
        Dim i As Integer = dat.Project.SwitchRobot(roboname)
        Console.WriteLine($"ロボット '{dat.Project.RobotName}' を選択しました")

        'Dim ck As Boolean = False
        'Dim jh As New RpaCui.JsonHandler(Of List(Of RpaUpdater))

        'Dim rrus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)
        '' ロボット名でフィルター
        'rrus = rrus.FindAll(
        '    Function(ru)
        '        Return ru.ReleaseTargets.Contains(roboname)
        '    End Function
        ')
        '' リリース日でソート
        'rrus.Sort(
        '    Function(before, after)
        '        Return (before.ReleaseDate < after.ReleaseDate)
        '    End Function
        ')

        'Dim srus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile)
        '' ロボット名でフィルター
        'srus = srus.FindAll(
        '    Function(ru)
        '        Return ru.ReleaseTargets.Contains(roboname)
        '    End Function
        ')
        '' リリース日でソート
        'srus.Sort(
        '    Function(before, after)
        '        Return (before.ReleaseDate < after.ReleaseDate)
        '    End Function
        ')

        'If rrus.Count > 0 Then
        '    If srus.Count = 0 Then
        '        ck = True
        '    Else
        '        If srus.Last.ReleaseDate < rrus.Last.ReleaseDate Then
        '            ck = True
        '        End If
        '    End If
        'End If

        'Dim targetrobots As New List(Of String)
        'Dim cmdstr As String = vbNullString
        'Dim cmdidf As String = vbNullString
        'If ck Then
        '    Console.WriteLine()
        '    Console.WriteLine($"最新のアップデートが適用されていません")
        '    For Each rru In rrus
        '        Console.WriteLine($"ReleaseDate  : {rru.ReleaseDate.ToString}")
        '        Console.WriteLine($"ReleaseTitle : {rru.ReleaseTitle}")
        '        Console.WriteLine($"ReleaseNote  : {rru.ReleaseNote}")
        '        Console.WriteLine()
        '        For Each targetrobot In rru.ReleaseTargets
        '            If Not targetrobots.Contains(targetrobot) Then
        '                targetrobots.Add(targetrobot)
        '            End If
        '        Next
        '    Next

        '    Console.WriteLine($"アップデートを適用するには、以下のコマンドを実行してください")
        '    cmdidf = dat.System.CommandDictionary.Where(
        '        Function(pair)
        '            Return (pair.Value.GetType.Name = (New UpdateRobotCommand).GetType.Name)
        '        End Function
        '    )(0).Key
        '    cmdstr = cmdidf
        '    For Each targetrobot In targetrobots
        '        cmdstr &= $" {targetrobot}"
        '    Next
        '    Console.WriteLine($"       '{cmdstr}'")
        '    Console.WriteLine($"または '{cmdidf}'")
        'End If

        Console.WriteLine()
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
        Me.ExecutableParameterCount = {0, 1}
    End Sub
End Class
