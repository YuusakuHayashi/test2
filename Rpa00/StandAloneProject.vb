Public Class StandAloneProject : Inherits RpaProjectBase(Of StandAloneProject)

    Public Overrides ReadOnly Property SystemJsonFileName As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property SystemArchDirectory As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property SystemArchType As Integer
        Get
            Return 2
        End Get
    End Property

    Public Overrides ReadOnly Property SystemArchTypeName As String
        Get
            Return Me.GetType.Name
        End Get
    End Property

    Public Overrides ReadOnly Property SystemProjectDirectory As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property SystemJsonChangeFileName As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property
End Class
