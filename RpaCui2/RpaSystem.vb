Imports System.IO

Module RpaSystem
    Private Delegate Function ExecuteDelegater(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer

    ' 機能はここに追加
    '---------------------------------------------------------------------------------------------'
    Private ReadOnly Property ExecuteHandler(ByVal cmd As String) As ExecuteDelegater
        Get
            Select Case cmd
                Case "update"
                    Return AddressOf UpdateProject
                Case "download"
                    Return AddressOf DownloadProject
                Case "project"
                    Return AddressOf SetProject
                Case "exit"
                    Return AddressOf RpaExit
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property
    '---------------------------------------------------------------------------------------------'


    ' プロジェクトのアップデート
    '---------------------------------------------------------------------------------------------'
    Private Function UpdateProject(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        Dim i = -1

        If rpa.IsRootProjectExists Then
            If trn.Parameters.Count > 0 Then
                If rpa.MyProjectObject.UpdatePackages.Contains > 0 Then
                End If
            End If
            Else
                Console.WriteLine("プロジェクト: '" & rpa.ProjectName & "' のルートプロジェクトが存在しません")
                Console.WriteLine("ダウンロード先が存在しません")
            End If
            Return i
    End Function
    '---------------------------------------------------------------------------------------------'

    ' プロジェクトのダウンロード
    '---------------------------------------------------------------------------------------------'
    Private Function DownloadProject(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        Dim i = -1

        If rpa.IsRootProjectExists Then
            If trn.Parameters.Count > 0 Then
                Call SelectedDownload(trn, rpa)
            Else
                Call AllDownload(trn, rpa, rpa.RootProjectDirectory, rpa.MyProjectDirectory)
            End If
        Else
            Console.WriteLine("プロジェクト: '" & rpa.ProjectName & "' のルートプロジェクトが存在しません")
            Console.WriteLine("ダウンロード先が存在しません")
        End If
        Return i
    End Function

    Private Sub SelectedDownload(ByRef trn As RpaTransaction, ByRef rpa As RpaProject)
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
                Console.WriteLine("File.Copy... src: " & src)
                Console.WriteLine("          => dst: " & dst)
            ElseIf Directory.Exists(src) Then
                If Directory.Exists(dst) Then
                    Console.WriteLine("既にディレクトリが存在します: " & dst)
                Else
                    Directory.CreateDirectory(dst)
                    Console.WriteLine("Directory.CreateDirectory... src: " & src)
                    Console.WriteLine("                          => dst: " & dst)
                End If
            End If
        Next
    End Sub

    Private Sub AllDownload(ByRef trn As RpaTransaction, ByRef rpa As RpaProject,
                                  ByVal src As String, ByVal dst As String)
        If Not Directory.Exists(dst) Then
            Directory.CreateDirectory(dst)
            Console.WriteLine("Directory.CreateDirectory... src: " & src)
            Console.WriteLine("                          => dst: " & dst)
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
                        Call AllDownload(trn, rpa, src2, dst2)
                    Else
                        File.Copy(src2, dst2, True)
                        Console.WriteLine("File.Copy... src: " & src2)
                        Console.WriteLine("          => dst: " & dst2)
                    End If
                End If
            End If
        Next
    End Sub
    '---------------------------------------------------------------------------------------------'


    ' プロジェクトの切り替え
    '---------------------------------------------------------------------------------------------'
    Private Function SetProject(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        Dim i = -1
        If trn.Parameters.Count = 0 Then
            i = 10
        Else
            rpa.ProjectName = trn.Parameters(0)
            If Not rpa.IsRootProjectExists Then
                Console.WriteLine("プロジェクト: '" & rpa.ProjectName & "' のルートプロジェクトが存在しません")
                Console.WriteLine("プロジェクト名が不適切な可能性があります")
            End If
            i = 0
        End If
        Return i
    End Function
    '---------------------------------------------------------------------------------------------'


    ' プロジェクト終了
    '---------------------------------------------------------------------------------------------'
    Private Function RpaExit(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        trn.ExitFlag = True
        Return 0
    End Function
    '---------------------------------------------------------------------------------------------'

    Public Sub Main(ByRef trn As RpaTransaction, ByRef rpa As RpaProject)
        Dim fnc = ExecuteHandler(trn.MainCommand)
        If fnc IsNot Nothing Then
            trn.ReturnCode = fnc(trn, rpa)
        End If
    End Sub
End Module
