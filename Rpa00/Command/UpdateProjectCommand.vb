Imports System.IO
Imports Rpa00

Public Class UpdateProjectCommand : Inherits RpaCommandBase
    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        If String.IsNullOrEmpty(dat.Project.RootDirectory) Then
            Console.WriteLine($"'RootDirectory' が設定されていません")
            Return False
        End If
        If Not Directory.Exists(dat.Project.RootDirectory) Then
            Console.WriteLine($"RootDirectory '{dat.Project.RootDirectory}' は存在しません")
            Return False
        End If
        If Not Directory.Exists(dat.Project.RootDllDirectory) Then
            Console.WriteLine($"RootDllDirectory '{dat.Project.RootDllDirectory}' は存在しません")
            Return False
        End If
        If dat.Project.RobotAliasDictionary.Count = 0 Then
            Console.WriteLine($"プロジェクトにはロボットが存在しません")
            Return False
        End If
        Return True
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim i As Integer = -1
        If dat.Transaction.Parameters.Count = 0 Then
            i = IntaractiveUpdate(dat)
        Else
            i = SelectedUpdate(dat)
        End If
        Return i
    End Function

    Private Function IntaractiveUpdate(ByRef dat As RpaDataWrapper) As Integer
        Dim i As Integer = -1
        Dim idxtext As String = vbNullString
        Console.WriteLine()
        Do
            idxtext = vbNullString
            Console.WriteLine($"処理Ｎｏを選択してください")
            Console.WriteLine($"0.  最新の更新を適用")
            Console.WriteLine($"1.  特定の更新を適用・特定の更新へダウングレード")
            Console.WriteLine($"9.  中止する")
            idxtext = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
            If idxtext = "0" Or idxtext = "1" Or idxtext = "2" Or idxtext = "9" Then
                Exit Do
            End If
        Loop Until False

        If idxtext = "9" Then
            Return 0
        End If

        Select Case idxtext
            Case "0"
                i = LatestUpdate(dat)
            Case "1"
                i = PointUpdate(dat)
        End Select

        Return i
    End Function

    ' IDを指定したアップデート
    ' 自動アップデートなどで使用
    Private Function SelectedUpdate(ByRef dat As RpaDataWrapper) As Integer
        Dim jh As New Rpa00.JsonHandler(Of List(Of RpaUpdater))
        Dim rrus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)

        If rrus.Count = 0 Then
            Console.WriteLine($"更新元が存在しません")
            Console.WriteLine()
            Return 0
        End If

        Dim id As String = dat.Transaction.Parameters(0)
        Dim rru As RpaUpdater
        rru = rrus.Find(
            Function(upd)
                Return upd.ReleaseId = id
            End Function
        )
        If rru Is Nothing Then
            Console.WriteLine($"指定した更新元 '{id}' が存在しません")
            Console.WriteLine()
            Return 0
        End If

        Dim srus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile)
        srus.Sort(
            Function(before, after)
                Return (before.ReleaseDate < after.ReleaseDate)
            End Function
        )
        If (srus.Last).ReleaseId = id Then
            Console.WriteLine($"指定した更新 '{id}' は適用されています")
            Console.WriteLine()
            Return 0
        End If

        Dim uru As RpaUpdater = rru
        ' 指定したアップデートを適用するロジック（共通化を検討）
        '-----------------------------------------------------------------------------------------'
        For Each src In Directory.GetFiles(uru.PackageDirectory)
            Dim dst As String = $"{RpaCui.SystemPointUpdateDllDirectory}\{Path.GetFileName(src)}"
            File.Copy(src, dst, True)
            Console.WriteLine($"ファイルをコピー  src: '{src}'")
            Console.WriteLine($"               => dst: '{dst}'")
        Next

        Dim srus2 As List(Of RpaUpdater) = srus.FindAll(
            Function(ru)
                Return ru.ReleaseDate >= uru.ReleaseDate
            End Function
        )
        srus2.Sort(
            Function(before, after)
                Return (before.ReleaseDate < after.ReleaseDate)
            End Function
        )
        For Each sru In srus2
            srus.Remove(sru)
        Next

        uru.UpdaterProcessId = dat.Initializer.ProcessId
        srus.Add(uru)
        jh.Save(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile, srus)
        Console.WriteLine($"アップデートを適用するためには、再起動を行う必要があります")
        Console.WriteLine()
        '-----------------------------------------------------------------------------------------'

        Return 0
    End Function

    ' 最新アップデート
    Private Function LatestUpdate(ByRef dat As RpaDataWrapper) As Integer

        Dim rdic As Dictionary(Of String, String) = dat.Project.RobotAliasDictionary
        Dim udic As Dictionary(Of String, IntranetClientServerProject.RpaUtility) = dat.Project.SystemUtilities
        Dim robots As List(Of String) = rdic.Keys.ToList()
        Dim utils As List(Of String) = udic.Keys.ToList()

        Dim robotname As String = dat.Project.RobotName
        Dim jh As New Rpa00.JsonHandler(Of List(Of RpaUpdater))
        Dim rrus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)
        Dim srus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile)

        Dim rrus2 As List(Of RpaUpdater)
        ' .netには部分集合を判定する方法がなさそう・・・
        rrus2 = rrus.FindAll(
            Function(rru)
                For Each rdep In rru.RobotDependencies
                    If Not robots.Contains(rdep) Then
                        Return False
                    End If
                Next
                Return True
            End Function
        )
        rrus2 = rrus2.FindAll(
            Function(rru)
                For Each udep In rru.UtilityDependencies
                    If Not utils.Contains(udep) Then
                        Return False
                    End If
                Next
                Return True
            End Function
        )
        rrus2.Sort(
            Function(before, after)
                Return (before.ReleaseDate < after.ReleaseDate)
            End Function
        )
        Dim srus2 As List(Of RpaUpdater) = srus
        srus2.Sort(
            Function(before, after)
                Return (before.ReleaseDate < after.ReleaseDate)
            End Function
        )

        Dim ck As Boolean = False
        If rrus2.Count > 0 Then

            If srus2.Count = 0 Then
                ck = True
            Else
                If rrus2.Last.ReleaseDate > srus2.Last.ReleaseDate Then
                    ck = True
                End If
            End If
        End If

        Dim uru As RpaUpdater
        If ck Then
            uru = rrus2.Last
            Console.WriteLine($"アップデートを適用")
            Console.WriteLine($"Release Title : {uru.ReleaseTitle}")
            Console.WriteLine($"Release Date  : {uru.ReleaseDate}")
            For Each src In Directory.GetFiles(uru.PackageDirectory)
                Dim dst As String = $"{RpaCui.SystemUpdateDllDirectory}\{Path.GetFileName(src)}"
                File.Copy(src, dst, True)
                Console.WriteLine($"ファイルをコピー  src: '{src}'")
                Console.WriteLine($"               => dst: '{dst}'")
            Next

            uru.UpdaterProcessId = dat.Initializer.ProcessId
            srus.Add(uru)
            jh.Save(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile, srus)
            Console.WriteLine($"アップデートを適用するためには、再起動を行う必要があります")
            Console.WriteLine()
        Else
            Console.WriteLine($"アップデートの適用対象は存在しませんでした")
            Console.WriteLine()
        End If

        Return 0
    End Function

    ' インタラクティブモードから特定時点のアップデートを適用
    Private Function PointUpdate(ByRef dat As Object) As Integer
        Dim jh As New Rpa00.JsonHandler(Of RpaUpdater)

        Dim rdic As Dictionary(Of String, String) = dat.Project.RobotAliasDictionary
        Dim udic As Dictionary(Of String, IntranetClientServerProject.RpaUtility) = dat.Project.SystemUtilities
        Dim robots As List(Of String) = rdic.Keys.ToList()
        Dim utils As List(Of String) = udic.Keys.ToList()

        Dim rrus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)
        Dim srus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile)

        Dim rrus2 As List(Of RpaUpdater)
        rrus2 = rrus.FindAll(
            Function(rru)
                For Each rdep In rru.RobotDependencies
                    If Not robots.Contains(rdep) Then
                        Return False
                    End If
                Next
                Return True
            End Function
        )
        rrus2 = rrus2.FindAll(
            Function(rru)
                For Each udep In rru.UtilityDependencies
                    If Not utils.Contains(udep) Then
                        Return False
                    End If
                Next
                Return True
            End Function
        )
        rrus2.Sort(
            Function(before, after)
                Return (before.ReleaseDate < after.ReleaseDate)
            End Function
        )
        For Each rru In rrus2
            Console.WriteLine($"Release Index : {rrus2.IndexOf(rru)}")
            Console.WriteLine($"Release Title : {rru.ReleaseTitle}")
            Console.WriteLine($"Release Date  : {rru.ReleaseDate}")
            Console.WriteLine($"Release Note  : {rru.ReleaseNote}")
            Console.WriteLine()
        Next

        Dim idx As Integer = -1
        Do
            idx = -1
            Dim idxtext As String = vbNullString
            Console.WriteLine($"アップデート／ダウングレードを適用する 'ReleaseIndex' を指定してください")
            idxtext = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()

            Dim idx2 As Integer = -1
            If IsNumeric(idxtext) Then
                idx2 = Integer.Parse(idxtext)
                If idx2 <= (rrus2.Count - 1) Then
                    idx = idx2
                End If
            End If
            If idx < 0 Then
                Console.WriteLine($"'ReleaseIndex' : {idxtext} は不正です")
                Console.WriteLine()
                Continue Do
            End If

            Dim yorn As String = vbNullString
            Do
                Console.WriteLine($"よろしいですか？  'ReleaseIndex' : {idxtext} (y/n)")
                yorn = dat.Transaction.ShowRpaIndicator(dat)
                If yorn = "y" Or yorn = "n" Then
                    Exit Do
                End If
            Loop Until False
            If yorn = "y" Then
                Exit Do
            End If
        Loop Until False

        Dim uru As RpaUpdater = rrus2(idx)
        ' 指定したアップデートを適用するロジック（共通化を検討）
        '-----------------------------------------------------------------------------------------'
        For Each src In Directory.GetFiles(uru.PackageDirectory)
            Dim dst As String = $"{RpaCui.SystemPointUpdateDllDirectory}\{Path.GetFileName(src)}"
            File.Copy(src, dst, True)
            Console.WriteLine($"ファイルをコピー  src: '{src}'")
            Console.WriteLine($"               => dst: '{dst}'")
        Next

        Dim srus2 As List(Of RpaUpdater) = srus.FindAll(
            Function(ru)
                Return ru.ReleaseDate >= uru.ReleaseDate
            End Function
        )
        srus2.Sort(
            Function(before, after)
                Return (before.ReleaseDate < after.ReleaseDate)
            End Function
        )
        For Each sru In srus2
            srus.Remove(sru)
        Next

        uru.UpdaterProcessId = dat.Initializer.ProcessId
        srus.Add(uru)

        jh.Save(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile, srus)
        Console.WriteLine($"アップデートを適用するためには、再起動を行う必要があります")
        Console.WriteLine()
        '-----------------------------------------------------------------------------------------'

        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
        Me.ExecutableParameterCount = {0, 1}
    End Sub
End Class
