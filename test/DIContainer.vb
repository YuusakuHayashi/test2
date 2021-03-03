Imports System.ComponentModel

Public Class DIContainer
    Public Shared ReadOnly Property [ShotokuCalculater] As ShotokuCalculater
        Get
            Return (New ShotokuCalculater)
        End Get
    End Property

    Public Shared ReadOnly Property ShotokuTaxUpdateChecker As Action(Of Model, PropertyChangedEventArgs)
        Get
            Return (AddressOf CheckUpdateShotokuTax)
        End Get
    End Property

    Public Shared Sub CheckUpdateShotokuTax(ByVal m As Model, ByVal e As PropertyChangedEventArgs)
        If e.PropertyName = "ShotokuTax" Then

        End If
    End Sub
End Class
