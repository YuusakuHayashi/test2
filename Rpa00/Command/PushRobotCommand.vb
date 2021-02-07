Imports System.IO

Public Class PushRobotCommand : Inherits RpaCommandBase
    ' Pushを禁止するファイル名
    Private _ProhibitedPushRobots As List(Of String)
    Private ReadOnly Property ProhibitedPushRobots As List(Of String)
        Get
            If Me._ProhibitedPushRobots Is Nothing Then
                Me._ProhibitedPushRobots = New List(Of String)
                Me._ProhibitedPushRobots.Add("RpaCui.exe")
            End If
            Return Me._ProhibitedPushRobots
        End Get
    End Property

    ' Pushを許すファイル拡張子
    Private _AllowPushExtensions As List(Of String)
    Private ReadOnly Property AllowPushExtensions As List(Of String)
        Get
            If Me._AllowPushExtensions Is Nothing Then
                Me._AllowPushExtensions = New List(Of String)
                Me._AllowPushExtensions.Add(".dll")
                Me._AllowPushExtensions.Add(".exe")
            End If
            Return Me._AllowPushExtensions
        End Get
    End Property

    Private _Updater As RpaUpdater
    Private Property Updater As RpaUpdater
        Get
            Return Me._Updater
        End Get
        Set(value As RpaUpdater)
            Me._Updater = value
        End Set
    End Property

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
        If String.IsNullOrEmpty(dat.Project.ReleaseRobotDirectory) Then
            Console.WriteLine($"'ReleaseRobotDirectory' が設定されていません")
            Return False
        End If
        If Not Directory.Exists(dat.Project.ReleaseRobotDirectory) Then
            Console.WriteLine($"ReleaseRobotDirectory '{dat.Project.ReleaseRobotDirectory}' は存在しません")
            Return False
        End If

        Return True
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim i As Integer = WriteReleaseInfo(dat)
        Dim j As Integer = AllPush(dat)

        Return 0
    End Function

    Private Function AllPush(ByRef dat As RpaDataWrapper) As Integer
        Dim srcdir As String = dat.Project.ReleaseRobotDirectory
        Dim dstdir As String = Me.Updater.PackageDirectory

        For Each src In Directory.GetFiles(srcdir)
            Dim ext As String = Path.GetExtension(src)
            Dim srcname As String = Path.GetFileName(src)
            If Me.AllowPushExtensions.Contains(ext) Then
                If Not Me.ProhibitedPushRobots.Contains(srcname) Then
                    Dim dst As String = $"{dstdir}\{srcname}"
                    File.Copy(src, dst, True)
                    Console.WriteLine($"ファイルをコピー  '{src}'")
                    Console.WriteLine($"               => '{dst}'")
                    Console.WriteLine()
                End If
            End If
        Next

        Return 0
    End Function

    Private Function WriteReleaseInfo(ByRef dat As RpaDataWrapper) As Integer
        Dim ru As New RpaUpdater With {
            .ReleaseDate = DateTime.Now.ToString("yyyyMMddHHmmss")
        }

        Console.WriteLine()
        Dim yorn As String = vbNullString
        Dim id As String = vbNullString
        Dim pdname As String = vbNullString
        Do
            yorn = vbNullString
            Console.WriteLine($"Please Input ReleaseTitle :")
            Dim ttl As String = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
            id = $"{ru.ReleaseDate}_{ttl}"
            pdname = $"{dat.Project.RootDllDirectory}\{id}"
            If Directory.Exists(pdname) Then
                Console.WriteLine($"Directory '{pdname}' is Already Exists,")
                Console.WriteLine($"Please Input Other ReleaseTitle.")
                Console.WriteLine()
                Continue Do
            End If
            Do
                Console.WriteLine($"ReleaseTitle : '{ttl}' ok? (y/n)")
                yorn = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn = "y" Or yorn = "n" Then
                    Exit Do
                End If
            Loop Until False
            ru.ReleaseTitle = ttl
            ru.ReleaseId = id
            ru.PackageDirectory = pdname
            If yorn = "y" Then
                Exit Do
            End If
        Loop Until False

        Dim yorn2 As String = vbNullString
        Do
            yorn2 = vbNullString
            Console.WriteLine($"Please Input ReleaseNote :")
            Dim note As String = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
            Do
                Console.WriteLine($"ReleaseNote : '{note}' ok? (y/n)")
                yorn2 = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn2 = "y" Or yorn2 = "n" Then
                    Exit Do
                End If
            Loop Until False
            ru.ReleaseNote = note
            If yorn2 = "y" Then
                Exit Do
            End If
        Loop Until False

        'Dim yorn4 As String = vbNullString
        'Dim yorn5 As String = vbNullString
        'Do
        '    yorn4 = vbNullString
        '    yorn5 = vbNullString
        '    Console.WriteLine($"Please Input ReleaseTargets :")
        '    Dim targetsstr As String = dat.Transaction.ShowRpaIndicator(dat)
        '    Dim targets() As String = targetsstr.Split(" "c)
        '    Console.WriteLine()
        '    Do
        '        Console.WriteLine($"ReleaseTarget : '{targetsstr}' add? (y/n)")
        '        yorn4 = dat.Transaction.ShowRpaIndicator(dat)
        '        Console.WriteLine()
        '        If yorn4 = "y" Or yorn4 = "n" Then
        '            Exit Do
        '        End If
        '    Loop Until False
        '    If yorn4 = "n" Then
        '        Continue Do
        '    End If
        '    For Each target In targets
        '        ru.ReleaseTargets.Add(target)
        '    Next
        '    Do
        '        Console.WriteLine($"continue? (y/n)")
        '        yorn5 = dat.Transaction.ShowRpaIndicator(dat)
        '        Console.WriteLine()
        '        If yorn5 = "y" Or yorn5 = "n" Then
        '            Exit Do
        '        End If
        '    Loop Until False
        '    If yorn5 = "n" Then
        '        Exit Do
        '    End If
        'Loop Until False

        ' ＤＬＬ依存関係の入力
        Do
            Dim yorn4 As String = vbNullString
            Dim yorn5 As String = vbNullString
            Dim yorn6 As String = vbNullString
            Console.WriteLine($"Please Input Robot Dependencies :")
            Dim depsstr As String = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
            If String.IsNullOrEmpty(depsstr) Then
                Do
                    Console.WriteLine($"No Dependencies? (y/n)")
                    yorn5 = dat.Transaction.ShowRpaIndicator(dat)
                    Console.WriteLine()
                    If yorn5 = "y" Or yorn5 = "n" Then
                        Exit Do
                    End If
                Loop Until False
                If yorn5 = "y" Then
                    Exit Do
                End If
            End If
            Dim deps() As String = depsstr.Split(" "c)
            Do
                Console.WriteLine($"Dependencies : '{depsstr}' add? (y/n)")
                yorn4 = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn4 = "y" Or yorn4 = "n" Then
                    Exit Do
                End If
            Loop Until False
            If yorn4 = "n" Then
                Continue Do
            End If
            For Each dep In deps
                ru.RobotDependencies.Add(dep)
            Next
            Do
                Console.WriteLine($"continue? (y/n)")
                yorn6 = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn6 = "y" Or yorn6 = "n" Then
                    Exit Do
                End If
            Loop Until False
            If yorn6 = "n" Then
                Exit Do
            End If
        Loop Until False

        ' ＵＴＩＬＩＴＹ依存関係の入力
        Do
            Dim yorn4 As String = vbNullString
            Dim yorn5 As String = vbNullString
            Dim yorn6 As String = vbNullString
            Console.WriteLine($"Please Input Utility Dependencies :")
            Dim depsstr As String = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
            If String.IsNullOrEmpty(depsstr) Then
                Do
                    Console.WriteLine($"No Dependencies? (y/n)")
                    yorn5 = dat.Transaction.ShowRpaIndicator(dat)
                    Console.WriteLine()
                    If yorn5 = "y" Or yorn5 = "n" Then
                        Exit Do
                    End If
                Loop Until False
                If yorn5 = "y" Then
                    Exit Do
                End If
            End If
            Dim deps() As String = depsstr.Split(" "c)
            Do
                Console.WriteLine($"Dependencies : '{depsstr}' add? (y/n)")
                yorn4 = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn4 = "y" Or yorn4 = "n" Then
                    Exit Do
                End If
            Loop Until False
            If yorn4 = "n" Then
                Continue Do
            End If
            For Each dep In deps
                ru.UtilityDependencies.Add(dep)
            Next
            Do
                Console.WriteLine($"continue? (y/n)")
                yorn6 = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn6 = "y" Or yorn6 = "n" Then
                    Exit Do
                End If
            Loop Until False
            If yorn6 = "n" Then
                Exit Do
            End If
        Loop Until False

        Dim yorn7 As String = vbNullString
        Do
            yorn7 = vbNullString
            Console.WriteLine($"Is this Critical? (y/n) :")
            Dim yorn8 As String = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
            Do
                Console.WriteLine($"IsCritical : '{yorn8}' ok? (y/n)")
                yorn7 = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn7 = "y" Or yorn7 = "n" Then
                    Exit Do
                End If
            Loop Until False
            ru.IsCritical = IIf(yorn8 = "y", True, False)
            If yorn7 = "y" Then
                Exit Do
            End If
        Loop Until False

        'Dim yorn9 As String = vbNullString
        'Do
        '    yorn9 = vbNullString
        '    Console.WriteLine($"UpdatedBindingCommand ? :")
        '    Dim ubc As String = dat.Transaction.ShowRpaIndicator(dat)
        '    Console.WriteLine()
        '    If String.IsNullOrEmpty(ubc) Then
        '        ubc = vbNullString
        '    End If
        '    Do
        '        Console.WriteLine($"UpdatedBindingCommand : '{ubc}' ok? (y/n)")
        '        yorn9 = dat.Transaction.ShowRpaIndicator(dat)
        '        Console.WriteLine()
        '        If yorn9 = "y" Or yorn9 = "n" Then
        '            Exit Do
        '        End If
        '    Loop Until False
        '    ru.UpdatedBindingCommand = ubc
        '    If yorn9 = "y" Then
        '        Exit Do
        '    End If
        'Loop Until False

        ' バインディングコマンドの入力
        Do
            Dim yorn4 As String = vbNullString
            Dim yorn5 As String = vbNullString
            Dim yorn6 As String = vbNullString
            Console.WriteLine($"Please Input Updated Binding Commands :")
            Dim ubcsstr As String = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
            If String.IsNullOrEmpty(ubcsstr) Then
                Do
                    Console.WriteLine($"No Updated Binding Commands? (y/n)")
                    yorn5 = dat.Transaction.ShowRpaIndicator(dat)
                    Console.WriteLine()
                    If yorn5 = "y" Or yorn5 = "n" Then
                        Exit Do
                    End If
                Loop Until False
                If yorn5 = "y" Then
                    Exit Do
                End If
            End If
            Dim ubcs() As String = ubcsstr.Split(" "c)
            Do
                Console.WriteLine($"Updated Binding Commands : '{ubcsstr}' ok? (y/n)")
                yorn4 = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn4 = "y" Or yorn4 = "n" Then
                    Exit Do
                End If
            Loop Until False
            If yorn4 = "n" Then
                Continue Do
            End If
            For Each ubc In ubcs
                ru.UpdatedBindingCommands.Add(ubc)
            Next
            Do
                Console.WriteLine($"continue? (y/n)")
                yorn6 = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn6 = "y" Or yorn6 = "n" Then
                    Exit Do
                End If
            Loop Until False
            If yorn6 = "n" Then
                Exit Do
            End If
        Loop Until False

        Directory.CreateDirectory(pdname)

        Dim jh As New Rpa00.JsonHandler(Of List(Of RpaUpdater))
        Dim [old] As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)
        [old].Add(ru)

        Call jh.Save(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile, [old])
        Console.WriteLine($"'{dat.Project.RootRobotsUpdateFile}' Updated!")
        Console.WriteLine()

        Me.Updater = ru

        Return 0
    End Function

    Sub New()
        Me.ExecutableParameterCount = {0, 0}
        Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
        Me.ExecutableUser = {"RootDeveloper"}
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
