Imports System.Reflection

Public Class SetupProjectCommand : Inherits RpaCommandBase
    Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        If dat.Project Is Nothing Then
            Console.WriteLine($"現在、プロジェクトに入っていません")
            Console.WriteLine()
            Return False
        End If
        Return True
    End Function

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        dat.Transaction.Modes.Add($"{Me.GetType.Name}")
        Console.WriteLine($"'{Me.GetType.Name}' を起動します")
        Console.WriteLine()

        Do
            Call dat.System.Main(dat)
        Loop Until dat.Transaction.ExitFlag

        dat.Transaction.ExitFlag = False
        dat.Transaction.Modes.Remove(dat.Transaction.Modes.Remove($"{Me.GetType.Name}"))

        Return 0
    End Function
End Class
