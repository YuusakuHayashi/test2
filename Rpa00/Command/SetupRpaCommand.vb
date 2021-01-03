Public Class SetupRpaCommand : Inherits RpaCommandBase
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
