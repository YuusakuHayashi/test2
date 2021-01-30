Imports System.IO
Imports Rpa00

Public Class UpdateRobotCommand : Inherits RpaCommandBase
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
        Dim idxtext As String = vbNullString
        Console.WriteLine()
        Do
            idxtext = vbNullString
            Console.WriteLine($"   処理を選択してください")
            Console.WriteLine($"0  最新までアップデート")
            Console.WriteLine($"1  特定時点へアップデート／ダウングレード")
            Console.WriteLine($"9  中止する")
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

    Private Function LatestUpdate(ByRef dat As Object) As Integer
        Dim robotname As String = dat.Project.RobotName
        Dim jh As New RpaCui.JsonHandler(Of List(Of RpaUpdater))
        Dim rrus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)
        Dim srus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile)

        Dim rrus2 As List(Of RpaUpdater) = rrus
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

        DIm ck As Boolean = False
        If rrus2.Count > 0
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

    Private Function PointUpdate(ByRef dat As Object) As Integer
        Dim jh As New RpaCui.JsonHandler(Of RpaUpdater)
        Dim rrus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)
        Dim srus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile)

        Dim rrus2 As List(Of RpaUpdater) = rrus
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
        srus.Add(uru)

        jh.Save(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile, srus)
        Console.WriteLine($"アップデートを適用するためには、再起動を行う必要があります")
        Console.WriteLine()

        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
        Me.ExecutableParameterCount = {0, 0}
    End Sub
End Class
