Imports System.IO

Public Class ShowRobotReadMeCommand : Inherits RpaCommandBase

    ' 面倒なので、アクセサは実装しない
    Private ReadMeFileName As String

    Private Function Check(ByRef dat As Object) As Integer
        Me.ReadMeFileName = vbNullString

        If File.Exists(dat.Project.MyRobotReadMeFileName) Then
            Me.ReadMeFileName = dat.Project.MyRobotReadMeFileName
        End If
        If String.IsNullOrEmpty(Me.ReadMeFileName) Then
            Console.WriteLine($"MyRobotReadMeFileName '{dat.Project.MyRobotReadMeFileName}' が存在しません")
            Return False
        End If
        Return True
    End Function

    Private Function Main(ByRef dat As Object) As Integer
        ' 空にすると再ロードする
        dat.Project.MyRobotReadMe = vbNullString
        Console.WriteLine(dat.Project.MyRobotReadMe)
        'Dim txt As String
        'Dim sr As StreamReader
        'Try
        '    sr = New System.IO.StreamReader(
        '        Me.ReadMeFileName, System.Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING))
        '    txt = sr.ReadToEnd()
        '    Console.WriteLine(txt)
        'Catch ex As Exception
        '    Console.WriteLine(ex.Message)
        'Finally
        '    If sr IsNot Nothing Then
        '        sr.Close()
        '        sr.Dispose()
        '    End If
        'End Try
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
        Me.ExecutableParameterCount = {0, 0}
    End Sub
End Class
