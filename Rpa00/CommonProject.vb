Public Class CommonProject : Inherits RpaProjectBase(Of CommonProject)

    Public Overrides ReadOnly Property SystemJsonFileName As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property SystemProjectDirectory As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides Sub CheckProject()
        Throw New NotImplementedException()
    End Sub
End Class
