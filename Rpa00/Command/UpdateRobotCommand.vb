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
        If Not Directory.Exists(dat.Project.RootRobotDirectory) Then
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
            i = AllUpdate(dat)
        Else
            i = SelectedUpdate(dat)
        End If
        Return i
    End Function

    Private Function AllUpdate(ByRef dat As Object) As Integer
        Dim srcdir As String = dat.Project.RootDllDirectory
        Dim dstdir As String = RpaCui.SystemUpdateDllDirectory
        Dim hit As Integer = 0

        For Each pair In dat.Project.RobotAliasDictionary
            Dim src As String = $"{srcdir}\{pair.Key}.dll"
            Dim dst As String = $"{dstdir}\{Path.GetFileName(src)}"
            If Not File.Exists(src) Then
                Console.WriteLine($"ファイル '{src}' は存在しません")
                Console.WriteLine()
                Continue For
            End If
            File.Copy(src, dst, True)
            Console.WriteLine($"ファイルをコピー  '{src}'")
            Console.WriteLine($"               => '{dst}'")
            Console.WriteLine()
            hit += 1
        Next

        If hit = 0 Then
            Console.WriteLine($"ディレクトリ '{srcdir}' に対象ファイルは存在しませんでした")
            Console.WriteLine()
        Else
            Console.WriteLine($"アップデートを適用するには、再起動してください")
            Console.WriteLine()
        End If
        Return 0
    End Function

    Private Function SelectedUpdate(ByRef dat As Object) As Integer
        Dim srcdir As String = dat.Project.RootDllDirectory
        Dim dstdir As String = RpaCui.SystemUpdateDllDirectory
        Dim hit As Integer = 0
        For Each para In dat.Transaction.Parameters
            Dim src As String = $"{srcdir}\{para}.dll"
            Dim ext As String = Path.GetExtension(src)
            Dim dst As String = $"{dstdir}\{Path.GetFileName(src)}"
            If Not dat.Project.RobotAliasDictionary.ContainsKey(para) Then
                Console.WriteLine($"指定のアップデート対象ロボット '{para}' はプロジェクトに存在しません")
                Console.WriteLine()
                Continue For
            End If
            If Not File.Exists(src) Then
                Console.WriteLine($"ファイル '{src}' は存在しません")
                Console.WriteLine()
                Continue For
            End If
            File.Copy(src, dst, True)
            Console.WriteLine($"ファイルをコピー  '{src}'")
            Console.WriteLine($"               => '{dst}'")
            Console.WriteLine()
            hit += 1
        Next
        If hit = 0 Then
            Console.WriteLine($"ディレクトリ '{srcdir}' に対象ファイルは存在しませんでした")
            Console.WriteLine()
        Else
            Console.WriteLine($"アップデートを適用するには、再起動してください")
            Console.WriteLine()
        End If
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
        Me.ExecutableParameterCount = {0, 999}
    End Sub
End Class
