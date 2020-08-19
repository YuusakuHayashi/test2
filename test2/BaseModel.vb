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

    '2020/08/18 Add .
    'Friend Sub [Set](Of T)(ByRef storage As T, newValue As T, <CallerMemberName> Optional propertyName As String = Nothing)
    '    If Object.Equals(storage, newValue) Then
    '        Return
    '    End If
    '    storage = newValue
    '    RaisePropertyChanged(propertyName)
    'End Sub
End Class
