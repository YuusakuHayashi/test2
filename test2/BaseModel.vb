Imports System.ComponentModel

Public MustInherit Class BaseModel(Of T)
    Inherits JsonHandler(Of T)
    Implements INotifyPropertyChanged

    'Public ReadOnly Property ClassName As String
    '    Get
    '        Return Me.GetType.Name
    '    End Get
    'End Property

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Public Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub
End Class
