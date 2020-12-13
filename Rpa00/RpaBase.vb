Public MustInherit Class RpaBase(Of T As {New}) : Inherits RpaCui2.JsonHandler(Of T)
    Public MustOverride Function SetupProjectObject(ByVal project As String) As Object
    Public MustOverride Function Main() As Integer

    Protected Rpa As RpaProject
    Protected Transaction As RpaTransaction

    Public Sub SetData(ByRef trn As RpaTransaction, ByRef rpa As RpaProject)
        Me.Rpa = rpa
        Me.Transaction = trn
    End Sub
End Class