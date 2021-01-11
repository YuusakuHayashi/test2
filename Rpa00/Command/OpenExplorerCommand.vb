Imports System.IO

Public Class OpenExplorerCommand : Inherits RpaCommandBase

    Private _OpenPath As String
    Private Property OpenPath As String
        Get
            Return Me._OpenPath
        End Get
        Set(value As String)
            Me._OpenPath = vbNullString
            If Not String.IsNullOrEmpty(value) Then
                If Directory.Exists(value) Then
                    Me._OpenPath = value
                End If
            End If
        End Set
    End Property

    Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        Me.OpenPath = vbNullString
        If dat.Project Is Nothing Then
            Me.OpenPath = RpaCui.SystemDirectory
        Else
            If Not String.IsNullOrEmpty(dat.Project.MyDirectory) And Directory.Exists(dat.Project.MyDirectory) Then
                Me.OpenPath = dat.Project.MyDirectory
            End If
            If Not String.IsNullOrEmpty(dat.Project.MyRobotDirectory) And Directory.Exists(dat.Project.MyRobotDirectory) Then
                Me.OpenPath = dat.Project.MyRobotDirectory
            End If
        End If

        If String.IsNullOrEmpty(Me.OpenPath) Then
            Console.WriteLine($"オープンするディレクトリがありません")
            Console.WriteLine()
            Return False
        End If

        Console.WriteLine()
        Return True
    End Function

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Diagnostics.Process.Start(Me.OpenPath)
        Return 0
    End Function
End Class
