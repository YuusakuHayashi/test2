Imports test2
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class RpaModel
    Public RootProjectName As String
    Public UserProjectName As String
    Public Index As Integer
    Public Status As Integer
    <JsonIgnore>
    Public IsViewAssigned As Boolean
    Public RunParameter As String
End Class
