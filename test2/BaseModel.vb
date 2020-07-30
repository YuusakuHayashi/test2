Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports Newtonsoft.Json
Imports System.IO

Delegate Sub mySub()
Delegate Function myBoolFunction() As Boolean
Delegate Function myStringFunction() As String
Delegate Sub mySubWithT(Of T)(ByVal t As T)

Public MustInherit Class BaseModel
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub
End Class
