Public MustInherit Class RpaBase(Of T As {New}) : Inherits JsonHandler(Of T)
    Public MustOverride Function SetupProjectObject(ByVal project As String) As Object
    'Public MustOverride Function Main(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
    Public MustOverride Function Main() As Integer
    Public Sub SetData(ByRef trn As RpaTransaction, ByRef rpa As RpaProject)
        Me.Rpa = rpa
        Me.Transaction = trn
    End Sub
    Protected Rpa As RpaProject
    Protected Transaction As RpaTransaction
End Class
