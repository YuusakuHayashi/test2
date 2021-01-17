Imports System.IO

Public Class CloneRobotCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        If dat.Transaction.Parameters.Count > 0 Then
            Call SelectedCopy(dat)
        Else
            Call AllCopy(dat, dat.Project.RootRobotDirectory, dat.Project.MyRobotDirectory)
        End If
        Console.WriteLine("コピーが完了しました")
        Console.WriteLine()

        Return 0
    End Function

    Private Sub SelectedCopy(ByRef dat As RpaDataWrapper)
        For Each param In dat.Transaction.Parameters
            Dim src = $"{dat.Project.RootRobotDirectory}\{param}"
            Dim dst = $"{dat.Project.MyRobotDirectory}\{param}"
            Dim dstp = Path.GetDirectoryName(dst)
            If Not Directory.Exists(dstp) Then
                Console.WriteLine("ディレクトリが存在しません: " & dstp)
                Continue For
            End If
            If File.Exists(src) Then
                File.Copy(src, dst, True)
                Console.WriteLine($"ファイルをコピー　 {src}")
                Console.WriteLine($"                => {dst}")
            ElseIf Directory.Exists(src) Then
                If Directory.Exists(dst) Then
                    Console.WriteLine($"既にディレクトリ '{dst}' は存在します")
                Else
                    Directory.CreateDirectory(dst)
                    Console.WriteLine($"ディレクトリを作成 {src}")
                    Console.WriteLine($"                => {dst}")
                End If
            End If
        Next
    End Sub

    Private Sub AllCopy(ByRef dat As RpaDataWrapper, ByVal src As String, ByVal dst As String)
        If Not Directory.Exists(dst) Then
            Directory.CreateDirectory(dst)
            Console.WriteLine($"ディレクトリを作成 {src}")
            Console.WriteLine($"                => {dst}")
        End If

        Dim fsis As FileSystemInfo()
        Dim sdi = New DirectoryInfo(src)
        Dim ddi = New DirectoryInfo(dst)
        fsis = sdi.GetFileSystemInfos
        For Each fsi In fsis
            Dim src2 = $"{fsi.FullName}"
            Dim dst2 = $"{ddi.FullName}\{fsi.Name}"
            If Not dat.Project.RootRobotIgnoreList.Contains(src2) Then
                If Not dat.Project.MyRobotIgnoreList.Contains(src2) Then
                    If ((fsi.Attributes And FileAttributes.Directory) = FileAttributes.Directory) Then
                        Call AllCopy(dat, src2, dst2)
                    Else
                        File.Copy(src2, dst2, True)
                        Console.WriteLine($"ファイルをコピー　 {src2}")
                        Console.WriteLine($"                => {dst2}")
                    End If
                End If
            End If
        Next
    End Sub

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
        Me.ExecutableParameterCount = {0, 999}
    End Sub
End Class
