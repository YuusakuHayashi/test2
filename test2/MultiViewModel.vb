Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.ComponentModel
Imports System.Collections.ObjectModel

Public Class MultiViewModel
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
    Public Const PROJECT_MENU_FRAME As String = "ProjectMenuView"

    Public Const NORMAL_VIEW As String = "Normal"
    Public Const TAB_VIEW As String = "Tab"

    Private _MainViewHeight As Double
    Public Property MainViewHeight As Double
        Get
            Return Me._MainViewHeight
        End Get
        Set(value As Double)
            Me._MainViewHeight = value
        End Set
    End Property

    Private _MainViewPreservedHeight As Double
    Public Property MainViewPreservedHeight As Double
        Get
            Return Me._MainViewPreservedHeight
        End Get
        Set(value As Double)
            Me._MainViewPreservedHeight = value
        End Set
    End Property

    Private _MainGridHeight As GridLength
    <JsonIgnore>
    Public Property MainGridHeight As GridLength
        Get
            Return Me._MainGridHeight
        End Get
        Set(value As GridLength)
            Me._MainGridHeight = value
            Me.MainViewHeight = value.Value
            RaisePropertyChanged("MainGridHeight")
        End Set
    End Property

    Private _ExplorerViewHeight As Double
    Public Property ExplorerViewHeight As Double
        Get
            Return Me._ExplorerViewHeight
        End Get
        Set(value As Double)
            Me._ExplorerViewHeight = value
        End Set
    End Property

    Private _ExplorerViewPreservedHeight As Double
    Public Property ExplorerViewPreservedHeight As Double
        Get
            Return Me._ExplorerViewPreservedHeight
        End Get
        Set(value As Double)
            Me._ExplorerViewPreservedHeight = value
        End Set
    End Property

    Private _ExplorerGridHeight As GridLength
    <JsonIgnore>
    Public Property ExplorerGridHeight As GridLength
        Get
            Return Me._ExplorerGridHeight
        End Get
        Set(value As GridLength)
            Me._ExplorerGridHeight = value
            Me.ExplorerViewHeight = value.Value
        End Set
    End Property

    Private _HistoryViewHeight As Double
    Public Property HistoryViewHeight As Double
        Get
            Return Me._HistoryViewHeight
        End Get
        Set(value As Double)
            Me._HistoryViewHeight = value
        End Set
    End Property

    Private _HistoryViewPreservedHeight As Double
    Public Property HistoryViewPreservedHeight As Double
        Get
            Return Me._HistoryViewPreservedHeight
        End Get
        Set(value As Double)
            Me._HistoryViewPreservedHeight = value
        End Set
    End Property

    Private _HistoryGridHeight As GridLength
    <JsonIgnore>
    Public Property HistoryGridHeight As GridLength
        Get
            Return Me._HistoryGridHeight
        End Get
        Set(value As GridLength)
            Me._HistoryGridHeight = value
            Me.HistoryViewHeight = value.Value
        End Set
    End Property

    Private _RightViewWidth As Double
    Public Property RightViewWidth As Double
        Get
            Return Me._RightViewWidth
        End Get
        Set(value As Double)
            Me._RightViewWidth = value
        End Set
    End Property

    Private _RightViewPreservedWidth As Double
    Public Property RightViewPreservedWidth As Double
        Get
            Return Me._RightViewPreservedWidth
        End Get
        Set(value As Double)
            Me._RightViewPreservedWidth = value
        End Set
    End Property

    Private _LeftViewWidth As Double
    Public Property LeftViewWidth As Double
        Get
            Return Me._LeftViewWidth
        End Get
        Set(value As Double)
            Me._LeftViewWidth = value
        End Set
    End Property

    Private _LeftViewPreservedWidth As Double
    Public Property LeftViewPreservedWidth As Double
        Get
            Return Me._LeftViewPreservedWidth
        End Get
        Set(value As Double)
            Me._LeftViewPreservedWidth = value
        End Set
    End Property

    Private _RightGridWidth As GridLength
    <JsonIgnore>
    Public Property RightGridWidth As GridLength
        Get
            Return Me._RightGridWidth
        End Get
        Set(value As GridLength)
            Me._RightGridWidth = value
            Me.RightViewWidth = value.Value
            RaisePropertyChanged("RightGridWidth")
        End Set
    End Property

    Private _LeftGridWidth As GridLength
    <JsonIgnore>
    Public Property LeftGridWidth As GridLength
        Get
            Return Me._LeftGridWidth
        End Get
        Set(value As GridLength)
            Me._LeftGridWidth = value
            Me.LeftViewWidth = value.Value
            RaisePropertyChanged("LeftGridWidth")
        End Set
    End Property

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

    Private _ProjectMenuViewContent As Object
    <JsonIgnore>
    Public Property ProjectMenuViewContent As Object
        Get
            Return Me._ProjectMenuViewContent
        End Get
        Set(value As Object)
            Me._ProjectMenuViewContent = value
            RaisePropertyChanged("ProjectMenuViewContent")
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
