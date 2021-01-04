Imports Rpa00
Imports System.IO

' IntranetClientServerProject専用コマンド

Public Class AttachRobotCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecutableProjectArchitectures As String()
        Get
            Return {(New IntranetClientServerProject).GetType.Name}
        End Get
    End Property

    Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
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
        Else
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

    Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
        Get
            Return {0, 1}
        End Get
    End Property

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
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
        dat.Project.RobotName = roboname
        Console.WriteLine($"プロジェクト '{dat.Project.RobotName}' を選択しました")
        Console.WriteLine()
        Return 0
    End Function
End Class
