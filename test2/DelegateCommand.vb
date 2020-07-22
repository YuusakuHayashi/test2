Public Class DelegateCommand : Implements ICommand

    Public Property CanExecuteDelegater As Func(Of Object, Boolean)
    Public Property ExecuteDelegater As Action(Of String)

    Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

    Public Sub RaiseCanExecuteChanged()
        RaiseEvent CanExecuteChanged(Me, Nothing)
    End Sub

    Public Sub Execute(parameter As Object) Implements ICommand.Execute
        Dim d = ExecuteDelegater
        If d <> Nothing Then
            d(parameter)
        End If
    End Sub

    Public Function CanExecute(ByVal parameter As Object) As Boolean Implements ICommand.CanExecute
        Dim d = CanExecuteDelegater
        Return IIf(d = Nothing, True, d(parameter))
    End Function
End Class
