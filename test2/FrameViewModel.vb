Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class FrameViewModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Public Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub

    ' GRID LENGTH
    '---------------------------------------------------------------'
    Private _MenuGridHeight As GridLength
    <JsonIgnore>
    Public Property MenuGridHeight As GridLength
        Get
            Return Me._MenuGridHeight
        End Get
        Set(value As GridLength)
            Me._MenuGridHeight = value
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
        End Set
    End Property

    Private _MainGridWidth As GridLength
    <JsonIgnore>
    Public Property MainGridWidth As GridLength
        Get
            Return Me._MainGridWidth
        End Get
        Set(value As GridLength)
            Me._MainGridWidth = value
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
        End Set
    End Property

    Private _LeftExplorerGridWidth As GridLength
    <JsonIgnore>
    Public Property LeftExplorerGridWidth As GridLength
        Get
            Return Me._LeftExplorerGridWidth
        End Get
        Set(value As GridLength)
            Me._LeftExplorerGridWidth = value
        End Set
    End Property

    Private _RightExplorerGridWidth As GridLength
    <JsonIgnore>
    Public Property RightExplorerGridWidth As GridLength
        Get
            Return Me._RightExplorerGridWidth
        End Get
        Set(value As GridLength)
            Me._RightExplorerGridWidth = value
        End Set
    End Property
    '---------------------------------------------------------------'


    ' VIEW LENGTH
    '---------------------------------------------------------------'
    Private _MenuViewHeight As Double
    Public Property MenuViewHeight As Double
        Get
            Return Me._MenuViewHeight
        End Get
        Set(value As Double)
            Me._MenuViewHeight = value
        End Set
    End Property

    Private _MainViewHeight As Double
    Public Property MainViewHeight As Double
        Get
            Return Me._MainViewHeight
        End Get
        Set(value As Double)
            Me._MainViewHeight = value
        End Set
    End Property

    Private _MainViewWidth As Double
    Public Property MainViewWidth As Double
        Get
            Return Me._MainViewWidth
        End Get
        Set(value As Double)
            Me._MainViewWidth = value
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

    Private _LeftExplorerViewWidth As Double
    Public Property LeftExplorerViewWidth As Double
        Get
            Return Me._LeftExplorerViewWidth
        End Get
        Set(value As Double)
            Me._LeftExplorerViewWidth = value
        End Set
    End Property

    Private _RightExplorerViewWidth As Double
    Public Property RightExplorerViewWidth As Double
        Get
            Return Me._RightExplorerViewWidth
        End Get
        Set(value As Double)
            Me._RightExplorerViewWidth = value
        End Set
    End Property
    '---------------------------------------------------------------'


    ' GRID SPLITTER LENGTH
    '---------------------------------------------------------------'
    Private _MenuMainGridSplitterHeight As GridLength
    Public Property MenuMainGridSplitterHeight As GridLength
        Get
            Return Me._MenuMainGridSplitterHeight
        End Get
        Set(value As GridLength)
            Me._MenuMainGridSplitterHeight = value
        End Set
    End Property

    Private _MainHistoryGridSplitterHeight As GridLength
    Public Property MainHistoryGridSplitterHeight As GridLength
        Get
            Return Me._MainHistoryGridSplitterHeight
        End Get
        Set(value As GridLength)
            Me._MainHistoryGridSplitterHeight = value
        End Set
    End Property

    Private _LeftExplorerMainGridSplitterWidth As GridLength
    Public Property LeftExplorerMainGridSplitterWidth As GridLength
        Get
            Return Me._LeftExplorerMainGridSplitterWidth
        End Get
        Set(value As GridLength)
            Me._LeftExplorerMainGridSplitterWidth = value
        End Set
    End Property

    Private _MainRightExplorerGridSplitterWidth As GridLength
    Public Property MainRightExplorerGridSplitterWidth As GridLength
        Get
            Return Me._MainRightExplorerGridSplitterWidth
        End Get
        Set(value As GridLength)
            Me._MainRightExplorerGridSplitterWidth = value
        End Set
    End Property
    '---------------------------------------------------------------'


    ' GRID SPLITTER ISENABLED
    '---------------------------------------------------------------'
    Private _IsMenuMainGridSplitterEnabled As Boolean
    Public Property IsMenuMainGridSplitterEnabled As Boolean
        Get
            Return Me._IsMenuMainGridSplitterEnabled
        End Get
        Set(value As Boolean)
            Me._IsMenuMainGridSplitterEnabled = value
        End Set
    End Property

    Private _IsMainHistoryGridSplitterEnabled As Boolean
    Public Property IsMainHistoryGridSplitterEnabled As Boolean
        Get
            Return Me._IsMainHistoryGridSplitterEnabled
        End Get
        Set(value As Boolean)
            Me._IsMainHistoryGridSplitterEnabled = value
        End Set
    End Property

    Private _IsLeftExplorerMainGridSplitterEnabled As Boolean
    Public Property IsLeftExplorerMainGridSplitterEnabled As Boolean
        Get
            Return Me._IsLeftExplorerMainGridSplitterEnabled
        End Get
        Set(value As Boolean)
            Me._IsLeftExplorerMainGridSplitterEnabled = value
        End Set
    End Property

    Private _IsMainRightExplorerGridSplitterEnabled As Boolean
    Public Property IsMainRightExplorerGridSplitterEnabled As Boolean
        Get
            Return Me._IsMainRightExplorerGridSplitterEnabled
        End Get
        Set(value As Boolean)
            Me._IsMainRightExplorerGridSplitterEnabled = value
        End Set
    End Property
    '---------------------------------------------------------------'


    ' CONTENT VISIBILITY
    '---------------------------------------------------------------'
    Private _MenuContentVisibility As Visibility
    Public Property MenuContentVisibility As Visibility
        Get
            Return Me._MenuContentVisibility
        End Get
        Set(value As Visibility)
            Me._MenuContentVisibility = value
        End Set
    End Property

    Private _MainContentVisibility As Visibility
    Public Property MainContentVisibility As Visibility
        Get
            Return Me._MainContentVisibility
        End Get
        Set(value As Visibility)
            Me._MainContentVisibility = value
        End Set
    End Property

    Private _LeftExplorerContentVisibility As Visibility
    Public Property LeftExplorerContentVisibility As Visibility
        Get
            Return Me._LeftExplorerContentVisibility
        End Get
        Set(value As Visibility)
            Me._LeftExplorerContentVisibility = value
        End Set
    End Property

    Private _RightExplorerContentVisibility As Visibility
    Public Property RightExplorerContentVisibility As Visibility
        Get
            Return Me._RightExplorerContentVisibility
        End Get
        Set(value As Visibility)
            Me._RightExplorerContentVisibility = value
        End Set
    End Property

    Private _HistoryContentVisibility As Visibility
    Public Property HistoryContentVisibility As Visibility
        Get
            Return Me._HistoryContentVisibility
        End Get
        Set(value As Visibility)
            Me._HistoryContentVisibility = value
        End Set
    End Property
    '---------------------------------------------------------------'


    ' CONTENT OBJECT
    '---------------------------------------------------------------'
    Private _MenuViewContent As ViewItemModel
    Public Property MenuViewContent As ViewItemModel
        Get
            Return Me._MenuViewContent
        End Get
        Set(value As ViewItemModel)
            Me._MenuViewContent = value
        End Set
    End Property

    Private _MainViewContent As ViewItemModel
    Public Property MainViewContent As ViewItemModel
        Get
            Return Me._MainViewContent
        End Get
        Set(value As ViewItemModel)
            Me._MainViewContent = value
        End Set
    End Property

    Private _LeftExplorerViewContent As ViewItemModel
    Public Property LeftExplorerViewContent As ViewItemModel
        Get
            Return Me._LeftExplorerViewContent
        End Get
        Set(value As ViewItemModel)
            Me._LeftExplorerViewContent = value
        End Set
    End Property

    Private _HistoryViewContent As ViewItemModel
    Public Property HistoryViewContent As ViewItemModel
        Get
            Return Me._HistoryViewContent
        End Get
        Set(value As ViewItemModel)
            Me._HistoryViewContent = value
        End Set
    End Property

    Private _RightExplorerViewContent As ViewItemModel
    Public Property RightExplorerViewContent As ViewItemModel
        Get
            Return Me._RightExplorerViewContent
        End Get
        Set(value As ViewItemModel)
            Me._RightExplorerViewContent = value
        End Set
    End Property
    '---------------------------------------------------------------'

    Public Sub OptimizeView()
        Me.MenuMainGridSplitterHeight = New GridLength(5.0)
        Me.MainHistoryGridSplitterHeight = New GridLength(5.0)
        Me.LeftExplorerMainGridSplitterWidth = New GridLength(5.0)
        Me.MainRightExplorerGridSplitterWidth = New GridLength(5.0)
        Me.IsMenuMainGridSplitterEnabled = True
        Me.IsLeftExplorerMainGridSplitterEnabled = True
        Me.IsMainRightExplorerGridSplitterEnabled = True
        Me.IsMainHistoryGridSplitterEnabled = True

        If Me.MenuViewContent Is Nothing Then
            Me.IsMenuMainGridSplitterEnabled = False
            Me.MenuMainGridSplitterHeight = New GridLength(0.0)
        End If
        If Me.LeftExplorerViewContent Is Nothing Then
            If Me.MainViewContent Is Nothing Then
                If Me.RightExplorerViewContent Is Nothing Then
                    Me.IsMenuMainGridSplitterEnabled = False
                    Me.MenuMainGridSplitterHeight = New GridLength(0.0)
                End If
            End If
        End If
        If Me.LeftExplorerViewContent Is Nothing Then
            Me.IsLeftExplorerMainGridSplitterEnabled = False
            Me.LeftExplorerMainGridSplitterWidth = New GridLength(0.0)
        End If
        If Me.RightExplorerViewContent Is Nothing Then
            Me.IsMainRightExplorerGridSplitterEnabled = False
            Me.MainRightExplorerGridSplitterWidth = New GridLength(0.0)
        End If
        If Me.MainViewContent Is Nothing Then
            If Me.HistoryViewContent Is Nothing Then
                Me.IsLeftExplorerMainGridSplitterEnabled = False
                Me.LeftExplorerMainGridSplitterWidth = New GridLength(0.0)
                Me.IsMainRightExplorerGridSplitterEnabled = False
                Me.MainRightExplorerGridSplitterWidth = New GridLength(0.0)
            End If
        End If
        If Me.MainViewContent Is Nothing Then
            Me.IsMainHistoryGridSplitterEnabled = False
            Me.MainHistoryGridSplitterHeight = New GridLength(0.0)
        End If
        If Me.HistoryViewContent Is Nothing Then
            Me.IsMainHistoryGridSplitterEnabled = False
            Me.MainHistoryGridSplitterHeight = New GridLength(0.0)
        End If
    End Sub

    Sub New()
    End Sub
End Class
