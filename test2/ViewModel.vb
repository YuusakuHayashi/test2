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

    Public Const MAIN_FRAME As String = "MainView"
    Public Const EXPLORER_FRAME As String = "ExplorerView"
    Public Const HISTORY_FRAME As String = "HistoryView"
    Public Const MENU_FRAME As String = "MenuView"

    Public Const NORMAL_VIEW As String = "Normal"
    Public Const TAB_VIEW As String = "Tab"

    Private _MainViewContent As Object
    <JsonIgnore>
    Public Property MainViewContent As Object
        Get
            Return Me._MainViewContent
        End Get
        Set(value As Object)
            Me._MainViewContent = value
            RaisePropertyChanged("MainViewContent")
        End Set
    End Property


    Private _ExplorerViewContent As Object
    <JsonIgnore>
    Public Property ExplorerViewContent As Object
        Get
            Return Me._ExplorerViewContent
        End Get
        Set(value As Object)
            Me._ExplorerViewContent = value
            RaisePropertyChanged("ExplorerViewContent")
        End Set
    End Property

    Private _HistoryViewContent As Object
    <JsonIgnore>
    Public Property HistoryViewContent As Object
        Get
            Return Me._HistoryViewContent
        End Get
        Set(value As Object)
            Me._HistoryViewContent = value
            RaisePropertyChanged("HistoryViewContent")
        End Set
    End Property

    Private _MenuViewContent As Object
    <JsonIgnore>
    Public Property MenuViewContent As Object
        Get
            Return Me._MenuViewContent
        End Get
        Set(value As Object)
            Me._MenuViewContent = value
            RaisePropertyChanged("MenuViewContent")
        End Set
    End Property


    ' コンテントディクショナリ関連 ---------------------------------------------------------------'
    ' コンテントディクショナリ
    Private _ContentDictionary As Dictionary(Of String, Dictionary(Of String, Object))
    <JsonIgnore>
    Public Property ContentDictionary As Dictionary(Of String, Dictionary(Of String, Object))
        Get
            If Me._ContentDictionary Is Nothing Then
                Me._ContentDictionary = New Dictionary(Of String, Dictionary(Of String, Object))
            End If
            Return Me._ContentDictionary
        End Get
        Set(value As Dictionary(Of String, Dictionary(Of String, Object)))
            Me._ContentDictionary = value
        End Set
    End Property

    ' --------------------------------------------------------------------------------------------'
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
    ' --------------------------------------------------------------------------------------------'


    ' タブコレクション関連
    ' タブディクショナリ
    ' --------------------------------------------------------------------------------------------'
    Private _TabsDictionary As Dictionary(Of String, TabViewModel)
    <JsonIgnore>
    Public Property TabsDictionary As Dictionary(Of String, TabViewModel)
        Get
            If Me._TabsDictionary Is Nothing Then
                Me._TabsDictionary = New Dictionary(Of String, TabViewModel)
            End If
            Return Me._TabsDictionary
        End Get
        Set(value As Dictionary(Of String, TabViewModel))
            Me._TabsDictionary = value
        End Set
    End Property

End Class
