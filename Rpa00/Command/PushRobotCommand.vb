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
        If Not Directory.Exists(dat.Project.RootRobotDirectory) Then
            Console.WriteLine($"RootRobotDirectory '{dat.Project.RootDllDirectory}' は存在しません")
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
        Dim i As Integer = -1
        If dat.Transaction.Parameters.Count = 0 Then
            i = AllPush(dat)
        Else
            i = SelectedPush(dat)
        End If
        Return i
    End Function

    Private Function AllPush(ByRef dat As RpaDataWrapper) As Integer
        Dim srcdir As String = dat.Project.ReleaseRobotDirectory
        Dim dstdir As String = dat.Project.RootDllDirectory
        Dim hit As Integer = 0

        For Each src In Directory.GetFiles(srcdir)
            Dim dst As String = $"{dstdir}\{Path.GetFileName(src)}"
            If Path.GetExtension(src) = ".dll" Then
                File.Copy(src, dst, True)
                Console.WriteLine($"ファイルをコピー  '{src}'")
                Console.WriteLine($"               => '{dst}'")
                Console.WriteLine()
                hit += 1
            End If
        Next

        If hit = 0 Then
            Console.WriteLine($"ディレクトリ '{srcdir}' に対象ファイルは存在しませんでした")
            Console.WriteLine()
        End If
        Return 0
    End Function

    Private Function SelectedPush(ByRef dat As RpaDataWrapper) As Integer
        Dim srcdir As String = dat.Project.ReleaseRobotDirectory
        Dim dstdir As String = dat.Project.RootDllDirectory
        Dim hit As Integer = 0

        For Each para In dat.Transaction.Parameters
            Dim src As String = $"{srcdir}\{para}"
            Dim dst As String = $"{dstdir}\{Path.GetFileName(src)}"
            If File.Exists(src) And (Path.GetExtension(src) = ".dll") Then
                File.Copy(src, dst, True)
                Console.WriteLine($"ファイルをコピー  '{src}'")
                Console.WriteLine($"               => '{dst}'")
                Console.WriteLine()
                hit += 1
            Else
                Console.WriteLine($"ファイル '{src}' は存在しません")
                Console.WriteLine()
            End If
        Next

        If hit = 0 Then
            Console.WriteLine($"ディレクトリ '{srcdir}' に対象ファイルは存在しませんでした")
            Console.WriteLine()
        End If
        Return 0
    End Function

    Sub New()
        Me.ExecutableParameterCount = {0, 999}
        Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
        Me.ExecutableUser = {"RootDeveloper"}
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
