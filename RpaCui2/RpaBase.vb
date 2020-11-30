Public MustInherit Class RpaBase(Of T As {New}) : Inherits JsonHandler(Of T)
    Public MustOverride Function SetupProjectObject(ByVal project As String) As Object
    Public MustOverride Function Main(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
End Class
