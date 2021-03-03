Imports System.Collections.Generic

Public Class Model
    Private _Calculater As Calculater
    Public Property [Calculater] As Calculater
        Get
            If Me._Calculater Is Nothing Then
                Me._Calculater = DIContainer.ShotokuCalculater
            End If
            Return Me._Calculater
        End Get
        Set(value As Calculater)
            Me._Calculater = value
        End Set
    End Property
End Class
