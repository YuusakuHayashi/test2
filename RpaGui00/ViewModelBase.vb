Imports System.ComponentModel
Imports System.Collections.Specialized
Imports Newtonsoft.Json

Public Class ViewModelBase(Of T As {New})
    Inherits Rpa00.JsonHandler(Of T)
    Implements INotifyPropertyChanged, INotifyCollectionChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Event CollectionChanged As NotifyCollectionChangedEventHandler Implements INotifyCollectionChanged.CollectionChanged

    Protected Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub

    Protected Sub RaiseCollectionChanged(ByVal PropertyName As String)
        RaiseEvent CollectionChanged(Me, New NotifyCollectionChangedEventArgs(PropertyName))
    End Sub

    ' !!! RpaDataWrapper型のdllオブジェクトがキャストされる
    ' 別プロジェクトのせいか、RpaDataWrapper型としては定義できなかった
    ' (するとキャスト時エラーになる)
    'Private _Data As Object
    '<JsonIgnore>
    'Public Property Data As Object
    '    Get
    '        Return Me._Data
    '    End Get
    '    Set(value As Object)
    '        Me._Data = value
    '        Call RaisePropertyChanged("Data")
    '    End Set
    'End Property

    ' よくよく考えると、JsonHandlerを継承した時点で、このロジックも実装している必要がある
    Public Overridable Sub Initialize()
    End Sub
End Class
