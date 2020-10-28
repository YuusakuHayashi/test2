Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ViewModel
    Inherits JsonHandler(Of ViewModel)
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Public Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub

    Private _SingleView As SingleViewModel
    <JsonIgnore>
    Public Property SingleView As SingleViewModel
        Get
            If Me._SingleView Is Nothing Then
                Me._SingleView = New SingleViewModel
            End If
            Return Me._SingleView
        End Get
        Set(value As SingleViewModel)
            Me._SingleView = value
            RaisePropertyChanged("SingleView")
        End Set
    End Property

    Private _MultiView As MultiViewModel
    <JsonIgnore>
    Public Property MultiView As MultiViewModel
        Get
            If Me._MultiView Is Nothing Then
                Me._MultiView = New MultiViewModel
            End If
            Return Me._MultiView
        End Get
        Set(value As MultiViewModel)
            Me._MultiView = value
            RaisePropertyChanged("MultiView")
        End Set
    End Property

    Private _Content As Object
    <JsonIgnore>
    Public Property Content As Object
        Get
            Return Me._Content
        End Get
        Set(value As Object)
            Me._Content = value
            RaisePropertyChanged("Content")
        End Set
    End Property

    Private _Views As ObservableCollection(Of ViewItemModel)
    Public Property Views As ObservableCollection(Of ViewItemModel)
        Get
            If Me._Views Is Nothing Then
                Me._Views = New ObservableCollection(Of ViewItemModel)
            End If
            Return Me._Views
        End Get
        Set(value As ObservableCollection(Of ViewItemModel))
            Me._Views = value
        End Set
    End Property

    Public Const SINGLE_VIEW As String = "Single"
    Public Const MULTI_VIEW As String = "Multi"
End Class
