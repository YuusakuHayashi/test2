Imports System.IO

Public Class ExitCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer

        If File.Exists(dat.Project.SystemJsonChangedFileName) Then
            Dim yorn As String = vbNullString
            Do
                yorn = vbNullString
                Console.WriteLine()
                Console.WriteLine($"Projectに変更があります。変更を保存しますか？ (y/n)")
                yorn = dat.Transaction.ShowRpaIndicator(dat)
            Loop Until yorn = "y" Or yorn = "n"
            If yorn = "y" Then
                RpaModule.Save(dat.Project.SystemJsonFileName, dat.Project, dat.Project.SystemJsonChangedFileName)
            Else
                File.Delete(dat.Project.SystemJsonChangedFileName)
            End If
        End If

        If File.Exists(RpaInitializer.SystemIniChangedFileName) Then
            Dim yorn2 As String = vbNullString
            Do
                yorn2 = vbNullString
                Console.WriteLine()
                Console.WriteLine($"Initializerに変更があります。変更を保存しますか？ (y/n)")
                yorn2 = dat.Transaction.ShowRpaIndicator(dat)
            Loop Until yorn2 = "y" Or yorn2 = "n"
            If yorn2 = "y" Then
                RpaModule.Save(RpaInitializer.SystemIniFileName, dat.Initializer, RpaInitializer.SystemIniChangedFileName)
            Else
                File.Delete(RpaInitializer.SystemIniChangedFileName)
            End If
        End If

        dat.Transaction.ExitFlag = True

        Console.WriteLine()
        Return 0
    End Function
End Class
