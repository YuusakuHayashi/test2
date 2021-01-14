Imports Rpa00
Imports System.IO

Public Class ShowMyRobotJsonCommand : Inherits RpaCommandBase
    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        If Not File.Exists(dat.Project.MyRobotJsonFileName) Then
            Console.WriteLine($"ファイル '{dat.Project.MyRobotJsonFileName}' が存在しません")
            Console.WriteLine()
            Return False
        End If
        Return True
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim txt As String = dat.Project.MyRobotObject.GetSerializedText(dat.Project.MyRobotObject)
        Console.WriteLine(txt)
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
    End Sub
End Class
