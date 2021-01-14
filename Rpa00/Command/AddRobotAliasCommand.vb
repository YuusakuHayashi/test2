Imports System.IO

Public Class AddRobotAliasCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
        Get
            Return {2, 2}
        End Get
    End Property

    Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        Dim [key] As String = dat.Transaction.Parameters(0)
        Dim [alias] As String = dat.Transaction.Parameters(1)
        If Not dat.Project.RobotAliasDictionary.ContainsKey([key]) Then
            Console.WriteLine($"ロボット '{[key]}' は登録簿にありません")
            Console.WriteLine()
            Return False
        Else
            Return True
        End If
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim oldrobo As String = dat.Project.RobotAlias
        Dim olddir As String = dat.Project.MyRobotDirectory

        Dim [key] As String = dat.Transaction.Parameters(0)
        Dim [alias] As String = dat.Transaction.Parameters(1)
        dat.Project.RobotAliasDictionary([key]) = [alias]

        ' 前のフォルダから移行する
        Dim yorn As String = vbNullString
        Dim newrobo As String = dat.Project.RobotAlias
        Dim newdir As String = dat.Project.MyRobotDirectory
        Do
            yorn = vbNullString
            Console.WriteLine()
            Console.WriteLine($"'{oldrobo}' 各ファイルを '{newrobo}' へ移行しますか？ (y/n)")
            yorn = dat.Transaction.ShowRpaIndicator(dat)
        Loop Until yorn = "y" Or yorn = "n"
        If yorn = "y" Then
            Call AllCopy(dat, olddir, newdir)
            Directory.Delete(olddir, True)
        End If

        Console.WriteLine($"ロボット '{[key]}' に '{[alias]}' を登録しました")
        Console.WriteLine()

        Return 0
    End Function

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
            If ((fsi.Attributes And FileAttributes.Directory) = FileAttributes.Directory) Then
                Call AllCopy(dat, src2, dst2)
            Else
                File.Copy(src2, dst2, True)
                Console.WriteLine($"ファイルをコピー　 {src2}")
                Console.WriteLine($"                => {dst2}")
            End If
        Next
    End Sub

    Sub New()
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
