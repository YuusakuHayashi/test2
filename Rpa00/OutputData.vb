Imports System.ComponentModel

'Public Class OutputData
Public Class OutputData : Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub

    Private _OutputText As String
    Public Property OutputText As String
        Get
            Return Me._OutputText
        End Get
        Set(value As String)
            Me._OutputText = value
            RaisePropertyChanged("OutputText")
        End Set
    End Property
End Class
