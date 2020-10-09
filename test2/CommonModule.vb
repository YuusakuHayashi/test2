Public Class CommonModule

    Public Delegate Function CheckModelProxy(Of T)(ByVal f As String) As Integer

    Private Function _CheckLoadModel(Of T)(ByVal f As String) As Integer
        Dim returnCode As Integer : returnCode = -1
        Dim ml As New ModelLoader(Of Nullable)
        Dim m As T

        Try
            m = ml.ModelLoad(Of T)(f)
            returnCode = 0
        Catch ex As Exception
            returnCode = 99
        End Try

        _CheckLoadModel = returnCode
    End Function

    Public Function CheckLoadModel(Of T)(ByVal f As String) As Integer
        Return _CheckLoadModel(Of T)(f)
    End Function
End Class
