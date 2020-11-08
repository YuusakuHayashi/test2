Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module CommonModule
    Public Function ConvertJObjToObj(Of T As {New})(ByVal jobj As Object) As T
        Dim obj As Object
        Select Case jobj.GetType
            Case (New Object).GetType
                obj = CType(jobj, T)
            Case (New JObject).GetType
                ' Ｊｓｏｎからロードした場合は、JObject型になっている
                obj = jobj.ToObject(Of T)
            Case (New T).GetType
                obj = jobj
            Case Else
                Throw New Exception("CommonModule.ConvertJObjToObj Error!!!")
        End Select
        ConvertJObjToObj = obj
    End Function
End Module
