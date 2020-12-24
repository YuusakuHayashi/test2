Imports System.IO

Public Class CloneProjectCommand : Inherits RpaCommandBase


    Public Overrides ReadOnly Property ExecutableProjectArchitectures As Integer()
        Get
            Return {
                RpaCodes.ProjectArchitecture.ClientServer,
                RpaCodes.ProjectArchitecture.IntranetClientServer
            }
        End Get
    End Property

    Public Overrides ReadOnly Property CanExecute(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableParameterCount() As Integer()
        Get
            Return {1, 1}
        End Get
    End Property

    Public Overrides ReadOnly Property ExecutableUserLevel() As Integer
        Get
            Return RpaCodes.ProjectUserLevel.User
        End Get
    End Property

    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        If trn.Parameters.Count > 0 Then
            Call SelectedCopy(trn, rpa)
        Else
            Call AllCopy(trn, rpa, rpa.RootProjectDirectory, rpa.MyProjectDirectory)
        End If
        Console.WriteLine("コピーが完了しました")

        Return 0
    End Function

    Private Sub SelectedCopy(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject)
        Dim src = vbNullString
        Dim dst = vbNullString
        Dim dstp = vbNullString
        For Each p In trn.Parameters
            src = rpa.RootProjectDirectory & "\" & p
            dst = rpa.MyProjectDirectory & "\" & p
            dstp = Path.GetDirectoryName(dst)
            If Not Directory.Exists(dstp) Then
                Console.WriteLine("ディレクトリが存在しません: " & dstp)
                Continue For
            End If
            If File.Exists(src) Then
                File.Copy(src, dst, True)
                Console.WriteLine($"Directory.CreateDirectory... src: {src}")
                Console.WriteLine($"                          => dst: {dst}")
            ElseIf Directory.Exists(src) Then
                If Directory.Exists(dst) Then
                    Console.WriteLine("既にディレクトリが存在します: " & dst)
                Else
                    Directory.CreateDirectory(dst)
                    Console.WriteLine($"Directory.CreateDirectory... src: {src}")
                    Console.WriteLine($"                          => dst: {dst}")
                End If
            End If
        Next
    End Sub

    Private Sub AllCopy(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject,
                                  ByVal src As String, ByVal dst As String)
        If Not Directory.Exists(dst) Then
            Directory.CreateDirectory(dst)
            Console.WriteLine($"Directory.CreateDirectory... src: {src}")
            Console.WriteLine($"                          => dst: {dst}")
        End If

        Dim fsis As FileSystemInfo()
        Dim sdi = New DirectoryInfo(src)
        Dim ddi = New DirectoryInfo(dst)
        Dim src2 = vbNullString
        Dim dst2 = vbNullString
        fsis = sdi.GetFileSystemInfos
        For Each fsi In fsis
            src2 = fsi.FullName
            dst2 = ddi.FullName & "\" & fsi.Name
            If Not rpa.RootProjectIgnoreList.Contains(src2) Then
                If Not rpa.MyProjectIgnoreList.Contains(src2) Then
                    If ((fsi.Attributes And FileAttributes.Directory) = FileAttributes.Directory) Then
                        Call AllCopy(trn, rpa, src2, dst2)
                    Else
                        File.Copy(src2, dst2, True)
                        Console.WriteLine("File.Copy... src: " & src2)
                        Console.WriteLine("          => dst: " & dst2)
                    End If
                End If
            End If
        Next
    End Sub
End Class
