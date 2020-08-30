Public Class FileManagerModel
    Private _LatestSettingFileName As String
    Public Property LatestSettingFileName As String
        Get
            Return Me._LatestSettingFileName
        End Get
        Set(value As String)
            Me._LatestSettingFileName = value
        End Set
    End Property
End Class
