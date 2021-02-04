Imports Rpa00

Public Class ClearScreenCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Console.Clear()
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecuteIfNoProject = True
    End Sub
End Class
