Public Class ExitCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        dat.Transaction.ExitFlag = True
        Console.WriteLine()
        Return 0
    End Function
End Class
