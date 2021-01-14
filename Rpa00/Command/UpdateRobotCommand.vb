Imports System.IO
Imports Rpa00

Public Class UpdateRobotCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecutableProjectArchitectures As String()
        Get
            Return {(New IntranetClientServerProject).GetType.Name}
        End Get
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
        If String.IsNullOrEmpty(dat.Project.RobotName) Then
            Console.WriteLine($"'RobotName' が設定されていません")
            Return False
        End If
        If Not Directory.Exists(dat.Project.RootRobotDirectory) Then
            Console.WriteLine($"RootRobotDirectory '{dat.Project.RootRobotDirectory}' は存在しません")
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
        Dim srcdir As String = dat.Project.RootRobotDllRepositoryDirectory
        Dim dstdir As String = RpaCui.SystemUpdateDllDirectory
        Dim hit As Integer = 0
        For Each src In Directory.GetFiles(srcdir)
            Dim ext As String = Path.GetExtension(src)
            Dim dst As String = $"{dstdir}\{Path.GetFileName(src)}"
            If File.Exists(src) Then
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
        Else
            Console.WriteLine($"アップデートを適用するには、再起動してください")
            Console.WriteLine()
        End If
        Return 0
    End Function

    Private Function SelectedUpdate(ByRef dat As Object) As Integer
        Dim srcdir As String = dat.Project.RootRobotDllRepositoryDirectory
        Dim dstdir As String = RpaCui.SystemUpdateDllDirectory
        Dim hit As Integer = 0
        For Each para In dat.Transaction.Parameters
            Dim src As String = $"{srcdir}\{para}"
            Dim ext As String = Path.GetExtension(src)
            Dim dst As String = $"{dstdir}\{para}"
            If File.Exists(src) Then
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
        Else
            Console.WriteLine($"アップデートを適用するには、再起動してください")
            Console.WriteLine()
        End If
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
    End Sub
End Class
