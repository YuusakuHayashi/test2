Imports System.ComponentModel

Public Class BaseViewModel
    Implements INotifyPropertyChanged

    '--- INortify -------------------------------------------------------------------------------------'
    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub
    '--------------------------------------------------------------------------------------------------'
End Class
