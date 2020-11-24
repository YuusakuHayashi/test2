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
            RaisePropertyChanged("MenuGridHeight")
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
            RaisePropertyChanged("MainGridHeight")
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
            RaisePropertyChanged("MainGridWidth")
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
            RaisePropertyChanged("HistoryGridHeight")
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
            RaisePropertyChanged("LeftExplorerGridWidth")
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
            RaisePropertyChanged("RightExplorerGridWidth")
        End Set
    End Property
    '---------------------------------------------------------------'


    ' VIEW LENGTH
    '---------------------------------------------------------------'
    Private _MenuViewHeightFlag As Boolean
    Public Property MenuViewHeightFlag As Boolean
        Get
            Return Me._MenuViewHeightFlag
        End Get
        Set(value As Boolean)
            Me._MenuViewHeightFlag = value
        End Set
    End Property

    Private _MenuViewHeight As Double
    Public Property MenuViewHeight As Double
        Get
            Return Me._MenuViewHeight
        End Get
        Set(value As Double)
            Me._MenuViewHeight = value
            If value <> 0.0 Then
                Me.MenuViewPreservedHeight = value
                Me.MenuViewHeightFlag = True
            End If
        End Set
    End Property

    Private _MenuViewPreservedHeight As Double
    Public Property MenuViewPreservedHeight As Double
        Get
            Return Me._MenuViewPreservedHeight
        End Get
        Set(value As Double)
            Me._MenuViewPreservedHeight = value
        End Set
    End Property
    '-------------------------------------------'

    Private _MainViewHeightFlag As Boolean
    Public Property MainViewHeightFlag As Boolean
        Get
            Return Me._MainViewHeightFlag
        End Get
        Set(value As Boolean)
            Me._MainViewHeightFlag = value
        End Set
    End Property

    Private _MainViewHeight As Double
    Public Property MainViewHeight As Double
        Get
            Return Me._MainViewHeight
        End Get
        Set(value As Double)
            Me._MainViewHeight = value
            If value <> 0.0 Then
                Me.MainViewPreservedHeight = value
                Me.MainViewHeightFlag = True
            End If
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
    '-------------------------------------------'

    Private _MainViewWidthFlag As Boolean
    Public Property MainViewWidthFlag As Boolean
        Get
            Return Me._MainViewWidthFlag
        End Get
        Set(value As Boolean)
            Me._MainViewWidthFlag = value
        End Set
    End Property

    Private _MainViewWidth As Double
    Public Property MainViewWidth As Double
        Get
            Return Me._MainViewWidth
        End Get
        Set(value As Double)
            Me._MainViewWidth = value
            If value <> 0.0 Then
                Me.MainViewPreservedWidth = value
                Me.MainViewWidthFlag = True
            End If
        End Set
    End Property

    Private _MainViewPreservedWidth As Double
    Public Property MainViewPreservedWidth As Double
        Get
            Return Me._MainViewPreservedWidth
        End Get
        Set(value As Double)
            Me._MainViewPreservedWidth = value
        End Set
    End Property
    '-------------------------------------------'

    Private _LeftExplorerViewWidthFlag As Boolean
    Public Property LeftExplorerViewWidthFlag As Boolean
        Get
            Return Me._LeftExplorerViewWidthFlag
        End Get
        Set(value As Boolean)
            Me._LeftExplorerViewWidthFlag = value
        End Set
    End Property

    Private _LeftExplorerViewWidth As Double
    Public Property LeftExplorerViewWidth As Double
        Get
            Return Me._LeftExplorerViewWidth
        End Get
        Set(value As Double)
            Me._LeftExplorerViewWidth = value
            If value <> 0.0 Then
                Me.LeftExplorerViewPreservedWidth = value
                Me.LeftExplorerViewWidthFlag = True
            End If
        End Set
    End Property

    Private _LeftExplorerViewPreservedWidth As Double
    Public Property LeftExplorerViewPreservedWidth As Double
        Get
            Return Me._LeftExplorerViewPreservedWidth
        End Get
        Set(value As Double)
            Me._LeftExplorerViewPreservedWidth = value
        End Set
    End Property
    '-------------------------------------------'

    Private _RightExplorerViewWidthFlag As Boolean
    Public Property RightExplorerViewWidthFlag As Boolean
        Get
            Return Me._RightExplorerViewWidthFlag
        End Get
        Set(value As Boolean)
            Me._RightExplorerViewWidthFlag = value
        End Set
    End Property

    Private _RightExplorerViewWidth As Double
    Public Property RightExplorerViewWidth As Double
        Get
            Return Me._RightExplorerViewWidth
        End Get
        Set(value As Double)
            Me._RightExplorerViewWidth = value
            If value <> 0.0 Then
                Me.RightExplorerViewPreservedWidth = value
                Me.RightExplorerViewWidthFlag = True
            End If
        End Set
    End Property

    Private _RightExplorerViewPreservedWidth As Double
    Public Property RightExplorerViewPreservedWidth As Double
        Get
            Return Me._RightExplorerViewPreservedWidth
        End Get
        Set(value As Double)
            Me._RightExplorerViewPreservedWidth = value
        End Set
    End Property
    '-------------------------------------------'

    Private _HistoryViewHeightFlag As Boolean
    Public Property HistoryViewHeightFlag As Boolean
        Get
            Return Me._HistoryViewHeightFlag
        End Get
        Set(value As Boolean)
            Me._HistoryViewHeightFlag = value
        End Set
    End Property

    Private _HistoryViewHeight As Double
    Public Property HistoryViewHeight As Double
        Get
            Return Me._HistoryViewHeight
        End Get
        Set(value As Double)
            Me._HistoryViewHeight = value
            If value <> 0.0 Then
                Me.HistoryViewPreservedHeight = value
                Me.HistoryViewHeightFlag = True
            End If
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
    '---------------------------------------------------------------'


    ' GRID SPLITTER LENGTH
    '---------------------------------------------------------------'
    Private _MenuMainGridSplitterHeight As Double
    Public Property MenuMainGridSplitterHeight As Double
        Get
            Return Me._MenuMainGridSplitterHeight
        End Get
        Set(value As Double)
            Me._MenuMainGridSplitterHeight = value
            RaisePropertyChanged("MenuMainGridSplitterHeight")
        End Set
    End Property

    Private _MainHistoryGridSplitterHeight As Double
    Public Property MainHistoryGridSplitterHeight As Double
        Get
            Return Me._MainHistoryGridSplitterHeight
        End Get
        Set(value As Double)
            Me._MainHistoryGridSplitterHeight = value
            RaisePropertyChanged("MainHistoryGridSplitterHeight")
        End Set
    End Property

    Private _LeftExplorerMainGridSplitterWidth As Double
    Public Property LeftExplorerMainGridSplitterWidth As Double
        Get
            Return Me._LeftExplorerMainGridSplitterWidth
        End Get
        Set(value As Double)
            Me._LeftExplorerMainGridSplitterWidth = value
            RaisePropertyChanged("LeftExplorerMainGridSplitterWidth")
        End Set
    End Property

    Private _MainRightExplorerGridSplitterWidth As Double
    Public Property MainRightExplorerGridSplitterWidth As Double
        Get
            Return Me._MainRightExplorerGridSplitterWidth
        End Get
        Set(value As Double)
            Me._MainRightExplorerGridSplitterWidth = value
            RaisePropertyChanged("MainRightExplorerGridSplitterWidth")
        End Set
    End Property
    '---------------------------------------------------------------'

    ' GRID ROW & COLUMN & ROW SPAN & COLUMN SPAN
    '---------------------------------------------------------------'
    Private _MenuRow As Integer
    Public Property MenuRow As Integer
        Get
            Return Me._MenuRow
        End Get
        Set(value As Integer)
            Me._MenuRow = value
            RaisePropertyChanged("MenuRow")
        End Set
    End Property

    Private _MenuColumn As Integer
    Public Property MenuColumn As Integer
        Get
            Return Me._MenuColumn
        End Get
        Set(value As Integer)
            Me._MenuColumn = value
            RaisePropertyChanged("MenuColumn")
        End Set
    End Property

    Private _MenuColumnSpan As Integer
    Public Property MenuColumnSpan As Integer
        Get
            Return Me._MenuColumnSpan
        End Get
        Set(value As Integer)
            Me._MenuColumnSpan = value
            RaisePropertyChanged("MenuColumnSpan")
        End Set
    End Property

    Private _LeftExplorerRow As Integer
    Public Property LeftExplorerRow As Integer
        Get
            Return Me._LeftExplorerRow
        End Get
        Set(value As Integer)
            Me._LeftExplorerRow = value
            RaisePropertyChanged("LeftExplorerRow")
        End Set
    End Property

    Private _LeftExplorerColumn As Integer
    Public Property LeftExplorerColumn As Integer
        Get
            Return Me._LeftExplorerColumn
        End Get
        Set(value As Integer)
            Me._LeftExplorerColumn = value
            RaisePropertyChanged("LeftExplorerColumn")
        End Set
    End Property

    Private _LeftExplorerRowSpan As Integer
    Public Property LeftExplorerRowSpan As Integer
        Get
            Return Me._LeftExplorerRowSpan
        End Get
        Set(value As Integer)
            Me._LeftExplorerRowSpan = value
            RaisePropertyChanged("LeftExplorerRowSpan")
        End Set
    End Property

    Private _MainRow As Integer
    Public Property MainRow As Integer
        Get
            Return Me._MainRow
        End Get
        Set(value As Integer)
            Me._MainRow = value
            RaisePropertyChanged("MainRow")
        End Set
    End Property

    Private _MainColumn As Integer
    Public Property MainColumn As Integer
        Get
            Return Me._MainColumn
        End Get
        Set(value As Integer)
            Me._MainColumn = value
            RaisePropertyChanged("MainColumn")
        End Set
    End Property

    Private _MainRowSpan As Integer
    Public Property MainRowSpan As Integer
        Get
            Return Me._MainRowSpan
        End Get
        Set(value As Integer)
            Me._MainRowSpan = value
            RaisePropertyChanged("MainRowSpan")
        End Set
    End Property

    Private _RightExplorerRow As Integer
    Public Property RightExplorerRow As Integer
        Get
            Return Me._RightExplorerRow
        End Get
        Set(value As Integer)
            Me._RightExplorerRow = value
            RaisePropertyChanged("RightExplorerRow")
        End Set
    End Property

    Private _RightExplorerColumn As Integer
    Public Property RightExplorerColumn As Integer
        Get
            Return Me._RightExplorerColumn
        End Get
        Set(value As Integer)
            Me._RightExplorerColumn = value
            RaisePropertyChanged("RightExplorerColumn")
        End Set
    End Property

    Private _RightExplorerRowSpan As Integer
    Public Property RightExplorerRowSpan As Integer
        Get
            Return Me._RightExplorerRowSpan
        End Get
        Set(value As Integer)
            Me._RightExplorerRowSpan = value
            RaisePropertyChanged("RightExplorerRowSpan")
        End Set
    End Property

    Private _HistoryRow As Integer
    Public Property HistoryRow As Integer
        Get
            Return Me._HistoryRow
        End Get
        Set(value As Integer)
            Me._HistoryRow = value
            RaisePropertyChanged("HistoryRow")
        End Set
    End Property

    Private _HistoryColumn As Integer
    Public Property HistoryColumn As Integer
        Get
            Return Me._HistoryColumn
        End Get
        Set(value As Integer)
            Me._HistoryColumn = value
            RaisePropertyChanged("HistoryColumn")
        End Set
    End Property
    '---------------------------------------------------------------'

    ' GRID SPLITTER RESIZE BEHAVIOR
    '---------------------------------------------------------------'
    Private _MenuMainSplitterResizeBehavior As Integer
    Public Property MenuMainSplitterResizeBehavior As Integer
        Get
            Return Me._MenuMainSplitterResizeBehavior
        End Get
        Set(value As Integer)
            Me._MenuMainSplitterResizeBehavior = value
            RaisePropertyChanged("MenuMainSplitterResizeBehavior")
        End Set
    End Property

    Private _MainHistorySplitterResizeBehavior As Integer
    Public Property MainHistorySplitterResizeBehavior As Integer
        Get
            Return Me._MainHistorySplitterResizeBehavior
        End Get
        Set(value As Integer)
            Me._MainHistorySplitterResizeBehavior = value
            RaisePropertyChanged("MainHistorySplitterResizeBehavior")
        End Set
    End Property

    Private _LeftExplorerMainSplitterResizeBehavior As Integer
    Public Property LeftExplorerMainSplitterResizeBehavior As Integer
        Get
            Return Me._LeftExplorerMainSplitterResizeBehavior
        End Get
        Set(value As Integer)
            Me._LeftExplorerMainSplitterResizeBehavior = value
            RaisePropertyChanged("LeftExplorerMainSplitterResizeBehavior")
        End Set
    End Property

    Private _MainRightExplorerSplitterResizeBehavior As Integer
    Public Property MainRightExplorerSplitterResizeBehavior As Integer
        Get
            Return Me._MainRightExplorerSplitterResizeBehavior
        End Get
        Set(value As Integer)
            Me._MainRightExplorerSplitterResizeBehavior = value
            RaisePropertyChanged("MainRightExplorerSplitterResizeBehavior")
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
            RaisePropertyChanged("IsMenuMainGridSplitterEnabled")
        End Set
    End Property

    Private _IsMainHistoryGridSplitterEnabled As Boolean
    Public Property IsMainHistoryGridSplitterEnabled As Boolean
        Get
            Return Me._IsMainHistoryGridSplitterEnabled
        End Get
        Set(value As Boolean)
            Me._IsMainHistoryGridSplitterEnabled = value
            RaisePropertyChanged("IsMainHistoryGridSplitterEnabled")
        End Set
    End Property

    Private _IsLeftExplorerMainGridSplitterEnabled As Boolean
    Public Property IsLeftExplorerMainGridSplitterEnabled As Boolean
        Get
            Return Me._IsLeftExplorerMainGridSplitterEnabled
        End Get
        Set(value As Boolean)
            Me._IsLeftExplorerMainGridSplitterEnabled = value
            RaisePropertyChanged("IsLeftExplorerMainGridSplitterEnabled")
        End Set
    End Property

    Private _IsMainRightExplorerGridSplitterEnabled As Boolean
    Public Property IsMainRightExplorerGridSplitterEnabled As Boolean
        Get
            Return Me._IsMainRightExplorerGridSplitterEnabled
        End Get
        Set(value As Boolean)
            Me._IsMainRightExplorerGridSplitterEnabled = value
            RaisePropertyChanged("IsMainRightExplorerGridSplitterEnabled")
        End Set
    End Property
    '---------------------------------------------------------------'


    ' CONTENT VISIBILITY
    '---------------------------------------------------------------'
    Private _MenuVisibility As Visibility
    Public Property MenuVisibility As Visibility
        Get
            Return Me._MenuVisibility
        End Get
        Set(value As Visibility)
            Me._MenuVisibility = value
            RaisePropertyChanged("MenuVisibility")
        End Set
    End Property

    Private _MainVisibility As Visibility
    Public Property MainVisibility As Visibility
        Get
            Return Me._MainVisibility
        End Get
        Set(value As Visibility)
            Me._MainVisibility = value
            RaisePropertyChanged("MainVisibility")
        End Set
    End Property

    Private _LeftExplorerVisibility As Visibility
    Public Property LeftExplorerVisibility As Visibility
        Get
            Return Me._LeftExplorerVisibility
        End Get
        Set(value As Visibility)
            Me._LeftExplorerVisibility = value
            RaisePropertyChanged("LeftExplorerVisibility")
        End Set
    End Property

    Private _RightExplorerVisibility As Visibility
    Public Property RightExplorerVisibility As Visibility
        Get
            Return Me._RightExplorerVisibility
        End Get
        Set(value As Visibility)
            Me._RightExplorerVisibility = value
            RaisePropertyChanged("RightExplorerVisibility")
        End Set
    End Property

    Private _HistoryVisibility As Visibility
    Public Property HistoryVisibility As Visibility
        Get
            Return Me._HistoryVisibility
        End Get
        Set(value As Visibility)
            Me._HistoryVisibility = value
            RaisePropertyChanged("HistoryExplorerVisibility")
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
            RaisePropertyChanged("MenuViewContent")
            'RaisePropertyChanged("MenuViewContent.Content")
        End Set
    End Property

    Private _MainViewContent As ViewItemModel
    Public Property MainViewContent As ViewItemModel
        Get
            Return Me._MainViewContent
        End Get
        Set(value As ViewItemModel)
            Me._MainViewContent = value
            RaisePropertyChanged("MainViewContent")
        End Set
    End Property

    Private _LeftExplorerViewContent As ViewItemModel
    Public Property LeftExplorerViewContent As ViewItemModel
        Get
            Return Me._LeftExplorerViewContent
        End Get
        Set(value As ViewItemModel)
            Me._LeftExplorerViewContent = value
            RaisePropertyChanged("LeftExplorerViewContent")
        End Set
    End Property

    Private _HistoryViewContent As ViewItemModel
    Public Property HistoryViewContent As ViewItemModel
        Get
            Return Me._HistoryViewContent
        End Get
        Set(value As ViewItemModel)
            Me._HistoryViewContent = value
            RaisePropertyChanged("HistoryViewContent")
        End Set
    End Property

    Private _RightExplorerViewContent As ViewItemModel
    Public Property RightExplorerViewContent As ViewItemModel
        Get
            Return Me._RightExplorerViewContent
        End Get
        Set(value As ViewItemModel)
            Me._RightExplorerViewContent = value
            RaisePropertyChanged("RightExplorerViewContent")
        End Set
    End Property
    '---------------------------------------------------------------'

    Public Sub OpenView(ByRef vim As ViewItemModel)
        If vim.Equals(Me.MenuViewContent) Then
        End If
        If vim.Equals(Me.MainViewContent) Then
        End If
        If vim.Equals(Me.LeftExplorerViewContent) Then
            Me.LeftExplorerGridWidth = GridLength.Auto
            'Me.MainGridWidth = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
            Me.IsLeftExplorerMainGridSplitterEnabled = True
            Me.LeftExplorerMainGridSplitterWidth = 5.0
        End If
        If vim.Equals(Me.RightExplorerViewContent) Then
            Me.RightExplorerGridWidth = GridLength.Auto
            'Me.RightExplorerGridWidth = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
            Me.IsMainRightExplorerGridSplitterEnabled = True
            Me.MainRightExplorerGridSplitterWidth = 5.0
        End If
        If vim.Equals(Me.HistoryViewContent) Then
            'Me.HistoryGridHeight = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
            Me.HistoryGridHeight = GridLength.Auto
            Me.IsMainHistoryGridSplitterEnabled = True
            Me.MainHistoryGridSplitterHeight = 5.0
        End If
    End Sub

    Public Sub CloseView(ByRef vim As ViewItemModel)
        If vim.Equals(Me.MenuViewContent) Then
        End If
        If vim.Equals(Me.MainViewContent) Then
        End If
        If vim.Equals(Me.LeftExplorerViewContent) Then
            'Me.MainGridWidth = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
            Me.IsLeftExplorerMainGridSplitterEnabled = False
            Me.LeftExplorerMainGridSplitterWidth = 0.0
        End If
        If vim.Equals(Me.RightExplorerViewContent) Then
            'Me.MainGridWidth = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
            Me.IsMainRightExplorerGridSplitterEnabled = False
            Me.MainRightExplorerGridSplitterWidth = 0.0
        End If
        If vim.Equals(Me.HistoryViewContent) Then
            Me.HistoryGridHeight = New GridLength(0.0)
            'Me.MainGridHeight = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
            Me.IsMainHistoryGridSplitterEnabled = False
            Me.MainHistoryGridSplitterHeight = 0.0
        End If
    End Sub

    Public Sub OptimizeView()
        Dim ck1, ck2, ck3, ck4, ck5, ck6, ck7, ck8, ck9, ck10, ck11, ck12, ck13, ck14

        'GRID SPLITTER RESIZE BEHAVIOR
        Me.MenuMainSplitterResizeBehavior = GridResizeBehavior.PreviousAndNext
        Me.MainHistorySplitterResizeBehavior = GridResizeBehavior.PreviousAndNext
        Me.LeftExplorerMainSplitterResizeBehavior = GridResizeBehavior.PreviousAndNext
        Me.MainRightExplorerSplitterResizeBehavior = GridResizeBehavior.PreviousAndNext

        ' GRID SPLITTER LENGTH, GRID SPLITTER ISENABLED
        Me.MenuMainGridSplitterHeight = 5.0
        Me.MainHistoryGridSplitterHeight = 5.0
        Me.LeftExplorerMainGridSplitterWidth = 5.0
        Me.MainRightExplorerGridSplitterWidth = 5.0
        Me.IsMenuMainGridSplitterEnabled = True
        Me.IsLeftExplorerMainGridSplitterEnabled = True
        Me.IsMainRightExplorerGridSplitterEnabled = True
        Me.IsMainHistoryGridSplitterEnabled = True

        ' ck1
        '-----------------------------------------------------------------------------------------'
        ck1 = True
        If Me.MenuViewContent Is Nothing Then
            ck1 = False
        Else
            If Me.MenuViewContent.Content Is Nothing Then
                ck1 = False
            End If
        End If
        If Not ck1 Then
            Me.IsMenuMainGridSplitterEnabled = False
            Me.MenuMainGridSplitterHeight = 0.0
        End If
        '-----------------------------------------------------------------------------------------'

        ' ck2 & ck3 & ck4
        '-----------------------------------------------------------------------------------------'
        ck2 = True
        If Me.LeftExplorerViewContent Is Nothing Then
            ck2 = False
        Else
            If Me.LeftExplorerViewContent.Content Is Nothing Then
                ck2 = False
            End If
        End If
        ck3 = True
        If Me.MainViewContent Is Nothing Then
            ck3 = False
        Else
            If Me.MainViewContent.Content Is Nothing Then
                ck3 = False
            End If
        End If
        ck4 = True
        If Me.RightExplorerViewContent Is Nothing Then
            ck4 = False
        Else
            If Me.RightExplorerViewContent.Content Is Nothing Then
                ck4 = False
            End If
        End If
        If (ck2 = False) And (ck3 = False) And (ck4 = False) Then
            Me.IsMenuMainGridSplitterEnabled = False
            Me.MenuMainGridSplitterHeight = 0.0
        End If
        '-----------------------------------------------------------------------------------------'

        ' ck5
        '-----------------------------------------------------------------------------------------'
        ck5 = True
        If Me.LeftExplorerViewContent Is Nothing Then
            ck5 = False
        Else
            If Me.LeftExplorerViewContent.Content Is Nothing Then
                ck5 = False
            End If
        End If
        If Not ck5 Then
            Me.IsLeftExplorerMainGridSplitterEnabled = False
            Me.LeftExplorerMainGridSplitterWidth = 0.0
        End If
        '-----------------------------------------------------------------------------------------'

        ' ck6
        '-----------------------------------------------------------------------------------------'
        ck6 = True
        If Me.RightExplorerViewContent Is Nothing Then
            ck6 = False
        Else
            If Me.RightExplorerViewContent.Content Is Nothing Then
                ck6 = False
            End If
        End If
        If Not ck6 Then
            Me.IsMainRightExplorerGridSplitterEnabled = False
            Me.MainRightExplorerGridSplitterWidth = 0.0
        End If
        '-----------------------------------------------------------------------------------------'

        ' ck7 & ck8
        '-----------------------------------------------------------------------------------------'
        ck7 = True
        If Me.MainViewContent Is Nothing Then
            ck7 = False
        Else
            If Me.MainViewContent.Content Is Nothing Then
                ck7 = False
            End If
        End If
        ck8 = True
        If Me.HistoryViewContent Is Nothing Then
            ck8 = False
        Else
            If Me.HistoryViewContent.Content Is Nothing Then
                ck8 = False
            End If
        End If
        If (ck7 = False) And (ck8 = False) Then
            Me.IsLeftExplorerMainGridSplitterEnabled = False
            Me.LeftExplorerMainGridSplitterWidth = 0.0
            Me.IsMainRightExplorerGridSplitterEnabled = False
            Me.MainRightExplorerGridSplitterWidth = 0.0
        End If
        '-----------------------------------------------------------------------------------------'

        ' ck9
        '-----------------------------------------------------------------------------------------'
        ck9 = True
        If Me.MainViewContent Is Nothing Then
            ck9 = False
        Else
            If Me.MainViewContent.Content Is Nothing Then
                ck9 = False
            End If
        End If
        If Not ck9 Then
            Me.IsMainHistoryGridSplitterEnabled = False
            Me.MainHistoryGridSplitterHeight = 0.0
        End If
        '-----------------------------------------------------------------------------------------'

        ' ck10
        '-----------------------------------------------------------------------------------------'
        ck10 = True
        If Me.HistoryViewContent Is Nothing Then
            ck10 = False
        Else
            If Me.HistoryViewContent.Content Is Nothing Then
                ck10 = False
            End If
        End If
        If Not ck10 Then
            Me.IsMainHistoryGridSplitterEnabled = False
            Me.MainHistoryGridSplitterHeight = 0.0
        End If
        '-----------------------------------------------------------------------------------------'

        ' GRID LENGTH
        'If Me.MenuViewHeightFlag Then
        '    Me.MenuGridHeight = New GridLength(Me.MenuViewPreservedHeight)
        'Else
        '    Me.MenuGridHeight = GridLength.Auto
        'End If
        'If Me.MainViewHeightFlag Then
        '    Me.MainGridHeight = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
        'Else
        '    Me.MainGridHeight = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
        'End If
        'If Me.HistoryViewHeightFlag Then
        '    Me.HistoryGridHeight = New GridLength(Me.HistoryViewPreservedHeight)
        'Else
        '    Me.HistoryGridHeight = GridLength.Auto
        'End If
        'If Me.LeftExplorerViewWidthFlag Then
        '    Me.LeftExplorerGridWidth = New GridLength(Me.LeftExplorerViewPreservedWidth)
        'Else
        '    Me.LeftExplorerGridWidth = GridLength.Auto
        'End If
        'If Me.MainViewWidthFlag Then
        '    Me.MainGridWidth = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
        'Else
        '    Me.MainGridWidth = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
        'End If
        'If Me.RightExplorerViewWidthFlag Then
        '    Me.RightExplorerGridWidth = New GridLength(Me.RightExplorerViewPreservedWidth)
        'Else
        '    Me.RightExplorerGridWidth = GridLength.Auto
        'End If
        If Me.MenuViewHeightFlag Then
            Me.MenuGridHeight = New GridLength(Me.MenuViewPreservedHeight)
        Else
            Me.MenuGridHeight = GridLength.Auto
        End If
        Me.MainGridHeight = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
        If Me.HistoryViewHeightFlag Then
            Me.HistoryGridHeight = New GridLength(Me.HistoryViewPreservedHeight)
        Else
            Me.HistoryGridHeight = GridLength.Auto
        End If
        If Me.LeftExplorerViewWidthFlag Then
            Me.LeftExplorerGridWidth = New GridLength(Me.LeftExplorerViewPreservedWidth)
        Else
            Me.LeftExplorerGridWidth = GridLength.Auto
        End If
        Me.MainGridWidth = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
        If Me.RightExplorerViewWidthFlag Then
            Me.RightExplorerGridWidth = New GridLength(Me.RightExplorerViewPreservedWidth)
        Else
            Me.RightExplorerGridWidth = GridLength.Auto
        End If
        'Me.MenuGridHeight = GridLength.Auto
        'Me.MainGridHeight = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
        'Me.HistoryGridHeight = GridLength.Auto
        'Me.LeftExplorerGridWidth = GridLength.Auto
        'Me.MainGridWidth = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
        'Me.RightExplorerGridWidth = GridLength.Auto

        ' ck11
        '-----------------------------------------------------------------------------------------'
        ck11 = True
        If Me.MenuViewContent Is Nothing Then
            ck11 = False
        Else
            If Me.MenuViewContent.Content Is Nothing Then
                ck11 = False
            End If
        End If
        If Not ck11 Then
            Me.MenuGridHeight = New GridLength(0.0, GridUnitType.Star)
        End If
        '-----------------------------------------------------------------------------------------'

        ' ck12
        '-----------------------------------------------------------------------------------------'
        ck12 = True
        If Me.LeftExplorerViewContent Is Nothing Then
            ck12 = False
        Else
            If Me.LeftExplorerViewContent.Content Is Nothing Then
                ck12 = False
            End If
        End If
        If Not ck12 Then
            Me.LeftExplorerGridWidth = New GridLength(0.0, GridUnitType.Star)
        End If
        '-----------------------------------------------------------------------------------------'

        ' ck13
        '-----------------------------------------------------------------------------------------'
        ck13 = True
        If Me.RightExplorerViewContent Is Nothing Then
            ck13 = False
        Else
            If Me.RightExplorerViewContent.Content Is Nothing Then
                ck13 = False
            End If
        End If
        If Not ck13 Then
            Me.RightExplorerGridWidth = New GridLength(0.0, GridUnitType.Star)
        End If
        '-----------------------------------------------------------------------------------------'

        ' ck14
        '-----------------------------------------------------------------------------------------'
        ck14 = True
        If Me.HistoryViewContent Is Nothing Then
            ck14 = False
        Else
            If Me.HistoryViewContent.Content Is Nothing Then
                ck14 = False
            End If
        End If
        If Not ck14 Then
            Me.HistoryGridHeight = New GridLength(0.0, GridUnitType.Star)
        End If
        '-----------------------------------------------------------------------------------------'

        ' GRID ROW, GRID COLUMN
        Me.MenuRow = 0 + 1
        Me.MenuColumn = 0 + 1
        Me.MenuColumnSpan = 3

        Me.LeftExplorerRow = 1 + 1
        Me.LeftExplorerRowSpan = 2
        Me.LeftExplorerColumn = 0 + 1

        Me.MainRow = 1 + 1
        Me.MainColumn = 1 + 1

        Me.RightExplorerRow = 1 + 1
        Me.RightExplorerRowSpan = 2
        Me.RightExplorerColumn = 2 + 1

        Me.HistoryRow = 2 + 1
        Me.HistoryColumn = 1 + 1

        ' TEST
        '-----------------------------------------------------------------------------------------'
        Console.WriteLine("LeftExplorerGridWidth  = " & Str(Me.LeftExplorerGridWidth.Value))
        Console.WriteLine("MainGridWidth          = " & Str(Me.MainGridWidth.Value))
        Console.WriteLine("RightExplorerGridWidth = " & Str(Me.RightExplorerGridWidth.Value))
        Console.WriteLine("MenuGridHeight         = " & Str(Me.MenuGridHeight.Value))
        Console.WriteLine("MainGridHeight         = " & Str(Me.MainGridHeight.Value))
        Console.WriteLine("HistoryGridHeight      = " & Str(Me.HistoryGridHeight.Value))
        Console.WriteLine("LeftExplorerMainGridSplitterWidth  = " & Str(Me.LeftExplorerMainGridSplitterWidth))
        Console.WriteLine("MainRightExplorerGridSplitterWidth = " & Str(Me.MainRightExplorerGridSplitterWidth))
        Console.WriteLine("MenuMainGridSplitterHeight         = " & Str(Me.MenuMainGridSplitterHeight))
        Console.WriteLine("MainHistoryGridSplitterHeight      = " & Str(Me.MainHistoryGridSplitterHeight))
        '-----------------------------------------------------------------------------------------'
    End Sub

    ' コマンドプロパティ（エクスプローラービューの右切り替え）
    '---------------------------------------------------------------------------------------------'
    Private _SwitchToRightCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property SwitchToRightCommand As ICommand
        Get
            If Me._SwitchToRightCommand Is Nothing Then
                Me._SwitchToRightCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _SwitchToRightCommandExecute,
                    .CanExecuteHandler = AddressOf _SwitchToRightCommandCanExecute
                }
                Return Me._SwitchToRightCommand
            Else
                Return Me._SwitchToRightCommand
            End If
        End Get
    End Property

    Private Sub _CheckSwitchToRightCommandEnabled()
        Dim b = True
        Me._SwitchToRightCommandEnableFlag = b
    End Sub

    Private __SwitchToRightCommandEnableFlag As Boolean
    <JsonIgnore>
    Public Property _SwitchToRightCommandEnableFlag As Boolean
        Get
            Return Me.__SwitchToRightCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__SwitchToRightCommandEnableFlag = value
            RaisePropertyChanged("_SwitchToRightCommandEnableFlag")
            CType(SwitchToRightCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _SwitchToRightCommandExecute(ByVal parameter As Object)
        If Me.RightExplorerViewContent Is Nothing Then
            Me.RightExplorerViewContent = New ViewItemModel With {
                .Name = "RightExplorer",
                .IsVisible = True
            }
        End If
        If Me.RightExplorerViewContent.Content Is Nothing Then
            Me.RightExplorerViewContent.Content = New TabViewModel
        End If
        For Each vt In Me.LeftExplorerViewContent.Content.ViewTabs
            Me.RightExplorerViewContent.Content.AddTab(vt)
        Next
        For Each child In Me.LeftExplorerViewContent.Children
            Me.RightExplorerViewContent.Children.Add(child)
        Next
        Call OpenView(Me.RightExplorerViewContent)
        Call CloseView(Me.LeftExplorerViewContent)

        ' 完全に消す
        Me.LeftExplorerViewContent = Nothing
    End Sub

    Private Function _SwitchToRightCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._SwitchToRightCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    ' コマンドプロパティ（エクスプローラービューの右切り替え）
    '---------------------------------------------------------------------------------------------'
    Private _SwitchToLeftCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property SwitchToLeftCommand As ICommand
        Get
            If Me._SwitchToLeftCommand Is Nothing Then
                Me._SwitchToLeftCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _SwitchToLeftCommandExecute,
                    .CanExecuteHandler = AddressOf _SwitchToLeftCommandCanExecute
                }
                Return Me._SwitchToLeftCommand
            Else
                Return Me._SwitchToLeftCommand
            End If
        End Get
    End Property

    Private Sub _CheckSwitchToLeftCommandEnabled()
        Dim b = True
        Me._SwitchToLeftCommandEnableFlag = b
    End Sub

    Private __SwitchToLeftCommandEnableFlag As Boolean
    <JsonIgnore>
    Public Property _SwitchToLeftCommandEnableFlag As Boolean
        Get
            Return Me.__SwitchToLeftCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__SwitchToLeftCommandEnableFlag = value
            RaisePropertyChanged("_SwitchToLeftCommandEnableFlag")
            CType(SwitchToLeftCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _SwitchToLeftCommandExecute(ByVal parameter As Object)
        If Me.LeftExplorerViewContent Is Nothing Then
            Me.LeftExplorerViewContent = New ViewItemModel With {
                .Name = "LeftExplorer",
                .IsVisible = True
            }
        End If
        If Me.LeftExplorerViewContent.Content Is Nothing Then
            Me.LeftExplorerViewContent.Content = New TabViewModel
        End If
        For Each vt In Me.RightExplorerViewContent.Content.ViewTabs
            Me.LeftExplorerViewContent.Content.AddTab(vt)
        Next
        For Each child In Me.RightExplorerViewContent.Children
            Me.LeftExplorerViewContent.Children.Add(child)
        Next
        Call OpenView(Me.LeftExplorerViewContent)
        Call CloseView(Me.RightExplorerViewContent)

        ' 完全に消す
        Me.RightExplorerViewContent = Nothing
    End Sub

    Private Function _SwitchToLeftCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._SwitchToLeftCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    Public Function DeepCopy()
        Return MemberwiseClone()
    End Function

    Sub New()
        Call _CheckSwitchToLeftCommandEnabled()
        Call _CheckSwitchToRightCommandEnabled()
    End Sub
End Class
