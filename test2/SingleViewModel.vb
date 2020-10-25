Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class SingleViewModel
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Public Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub

    Private _Content As Object
    <JsonIgnore>
    Public Property Content As Object
        Get
            If Me._Content Is Nothing Then
                Me._Content = New Object
            End If
            Return Me._Content
        End Get
        Set(value As Object)
            Me._Content = value
            RaisePropertyChanged("Content")
        End Set
    End Property
End Class
