Public MustInherit Class RpaBase(Of T As {New}) : Inherits JsonHandler(Of T)
    Public MustOverride Function SetupProjectObject(ByVal project As String) As Object
    Public MustOverride Function Execute() As Integer
    Public MustOverride Function CanExecute() As Boolean

    Protected Rpa As IntranetClientServerProject
    Protected Transaction As RpaTransaction

    Public Sub SetData(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject)
        Me.Rpa = rpa
        Me.Transaction = trn
    End Sub
End Class
