Public MustInherit Class RpaBase(Of T As {New}) : Inherits JsonHandler(Of T)
    Public MustOverride Function SetupProjectObject(ByVal project As String) As Object
    Public MustOverride Function Execute(ByRef dat As Object) As Integer
    Public MustOverride Function CanExecute(ByRef dat As Object) As Boolean

    Public Data As Object
    'Protected Transaction As RpaTransaction

    'Public Sub SetData(ByRef dat As RpaDataWrapper)
    '    Me.Rpa = dat.Project
    '    Me.Transaction = dat.Transaction
    'End Sub
End Class
