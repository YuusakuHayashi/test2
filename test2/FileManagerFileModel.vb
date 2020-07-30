Public Class FileManagerFileModel
    Private _ConfigFileName As String
    Public Property ConfigFileName As String
        Get
            Return _ConfigFileName
        End Get
        Set(value As String)
            _ConfigFileName = value
            'RaisePropertyChanged("ConfigFileName")
        End Set
    End Property

    Sub New()
        'Dim mp As New ProjectModel
        'Me.ConfigFileName = mp.LoadFileManagerFile.ConfigFileName
    End Sub
End Class
