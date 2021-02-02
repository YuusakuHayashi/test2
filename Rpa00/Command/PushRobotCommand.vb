Imports System.IO

Public Class PushRobotCommand : Inherits RpaCommandBase
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

        Dim ck As Boolean = False
        For Each fil In Directory.GetFiles(dat.Project.ReleaseRobotDirectory)
            If Path.GetExtension(fil) = ".dll" Then
                ck = True
                Exit For
            End If
        Next
        If Not ck Then
            Console.WriteLine($"ディレクトリ '{dat.Project.ReleaseRobotDirectory}' には、アップデート可能なファイルが存在しませんでした")
            Return False
        End If

        Return True
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim i As Integer = -1

        Dim pdname As String = WriteReleaseInfo(dat)
        i = AllPush(dat, pdname)

        Return i
    End Function

    Private Function AllPush(ByRef dat As RpaDataWrapper, ByVal pdir As String) As Integer
        Dim srcdir As String = dat.Project.ReleaseRobotDirectory
        Dim dstdir As String = pdir

        For Each src In Directory.GetFiles(srcdir)
            Dim dst As String = $"{dstdir}\{Path.GetFileName(src)}"
            If Path.GetExtension(src) = ".dll" Then
                File.Copy(src, dst, True)
                Console.WriteLine($"ファイルをコピー  '{src}'")
                Console.WriteLine($"               => '{dst}'")
                Console.WriteLine()
            End If
        Next

        Return 0
    End Function

    'Private Function SelectedPush(ByRef dat As RpaDataWrapper, ByVal pdir As String) As Integer
    '    Dim srcdir As String = dat.Project.ReleaseRobotDirectory
    '    Dim dstdir As String = pdir

    '    For Each para In dat.Transaction.Parameters
    '        Dim src As String = $"{srcdir}\{para}"
    '        Dim dst As String = $"{dstdir}\{Path.GetFileName(src)}"
    '        If Path.GetExtension(src) = ".dll" Then
    '            File.Copy(src, dst, True)
    '            Console.WriteLine($"ファイルをコピー  '{src}'")
    '            Console.WriteLine($"               => '{dst}'")
    '            Console.WriteLine()
    '        End If
    '    Next

    '    Return 0
    'End Function

    Private Function WriteReleaseInfo(ByRef dat As RpaDataWrapper) As String
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

        Dim yorn6 As String = vbNullString
        Do
            yorn6 = vbNullString
            Console.WriteLine($"Is this Critical? (y/n) :")
            Dim yorn7 As String = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
            Do
                Console.WriteLine($"IsCritical : '{yorn7}' ok? (y/n)")
                yorn6 = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
                If yorn6 = "y" Or yorn6 = "n" Then
                    Exit Do
                End If
            Loop Until False
            ru.IsCritical = IIf(yorn7 = "y", True, False)
            If yorn6 = "y" Then
                Exit Do
            End If
        Loop Until False

        Directory.CreateDirectory(pdname)

        Dim jh As New RpaCui.JsonHandler(Of List(Of RpaUpdater))
        Dim [old] As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)
        [old].Add(ru)

        Call jh.Save(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile, [old])
        Console.WriteLine($"'{dat.Project.RootRobotsUpdateFile}' Updated!")
        Console.WriteLine()

        Return pdname
    End Function

    Sub New()
        Me.ExecutableParameterCount = {0, 0}
        Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
        Me.ExecutableUser = {"RootDeveloper"}
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
