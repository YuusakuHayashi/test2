Imports Newtonsoft.Json

Public MustInherit Class RpaBase(Of T As {New}) : Inherits JsonHandler(Of T)
    Public MustOverride Function SetupProjectObject(ByVal project As String) As Object
    Public MustOverride Function Execute(ByRef dat As Object) As Integer
    Public MustOverride Function CanExecute(ByRef dat As Object) As Boolean

    <JsonIgnore>
    Public Data As Object
End Class
