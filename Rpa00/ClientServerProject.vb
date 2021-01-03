Public Class ClientServerProject : Inherits RpaProjectBase(Of ClientServerProject)

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
            Return 3
        End Get
    End Property

    Public Overrides ReadOnly Property SystemArchTypeName As String
        Get
            Return Me.GetType.Name
        End Get
    End Property

    Public Overrides ReadOnly Property SystemTempJsonFileName As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property SystemProjectDirectory As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides Sub BeginTransaction()
        Throw New NotImplementedException()
    End Sub

    Public Overrides Function TransactionRollBack() As RpaProjectBase(Of ClientServerProject)
        Throw New NotImplementedException()
    End Function
End Class
