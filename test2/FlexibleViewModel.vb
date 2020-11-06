Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class FlexibleViewModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Public Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub

    Private _ContentViewHeight As Double
    Public Property ContentViewHeight As Double
        Get
            Return Me._ContentViewHeight
        End Get
        Set(value As Double)
            Me._ContentViewHeight = value
        End Set
    End Property

    Private _ContentViewPreservedHeight As Double
    Public Property ContentViewPreservedHeight As Double
        Get
            Return Me._ContentViewPreservedHeight
        End Get
        Set(value As Double)
            Me._ContentViewPreservedHeight = value
        End Set
    End Property

    Private _ContentGridHeight As GridLength
    <JsonIgnore>
    Public Property ContentGridHeight As GridLength
        Get
            Return Me._ContentGridHeight
        End Get
        Set(value As GridLength)
            Me._ContentGridHeight = value
            'Me.ContentViewHeight = value.Value
        End Set
    End Property

    Private _ContentViewWidth As Double
    Public Property ContentViewWidth As Double
        Get
            Return Me._ContentViewWidth
        End Get
        Set(value As Double)
            Me._ContentViewWidth = value
        End Set
    End Property

    Private _ContentViewPreservedWidth As Double
    Public Property ContentViewPreservedWidth As Double
        Get
            Return Me._ContentViewPreservedWidth
        End Get
        Set(value As Double)
            Me._ContentViewPreservedWidth = value
        End Set
    End Property

    Private _ContentGridWidth As GridLength
    <JsonIgnore>
    Public Property ContentGridWidth As GridLength
        Get
            Return Me._ContentGridWidth
        End Get
        Set(value As GridLength)
            Me._ContentGridWidth = value
            'Me.ContentViewWidth = value.Value
        End Set
    End Property

    Private _BottomViewHeight As Double
    Public Property BottomViewHeight As Double
        Get
            Return Me._BottomViewHeight
        End Get
        Set(value As Double)
            Me._BottomViewHeight = value
        End Set
    End Property

    Private _BottomViewPreservedHeight As Double
    Public Property BottomViewPreservedHeight As Double
        Get
            Return Me._BottomViewPreservedHeight
        End Get
        Set(value As Double)
            Me._BottomViewPreservedHeight = value
        End Set
    End Property

    Private _BottomGridHeight As GridLength
    <JsonIgnore>
    Public Property BottomGridHeight As GridLength
        Get
            Return Me._BottomGridHeight
        End Get
        Set(value As GridLength)
            Me._BottomGridHeight = value
            Me.BottomViewHeight = value.Value
            RaisePropertyChanged("BottomGridHeight")
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

    Private _HorizontalSplitterWidth As GridLength
    <JsonIgnore>
    Public Property HorizontalSplitterWidth As GridLength
        Get
            Return Me._HorizontalSplitterWidth
        End Get
        Set(value As GridLength)
            Me._HorizontalSplitterWidth = value
            RaisePropertyChanged("HorizontalSplitterWidth")
        End Set
    End Property

    Private _VerticalSplitterHeight As GridLength
    <JsonIgnore>
    Public Property VerticalSplitterHeight As GridLength
        Get
            Return Me._VerticalSplitterHeight
        End Get
        Set(value As GridLength)
            Me._VerticalSplitterHeight = value
            RaisePropertyChanged("VerticalSplitterHeight")
        End Set
    End Property

    Private _MainContent As Object
    <JsonIgnore>
    Public Property MainContent As Object
        Get
            Return Me._MainContent
        End Get
        Set(value As Object)
            Me._MainContent = value
            Call _OptimizeDynamicView()
            RaisePropertyChanged("MainContent")
        End Set
    End Property

    Private _MainViewContent As ViewItemModel
    Public Property MainViewContent As ViewItemModel
        Get
            Return Me._MainViewContent
        End Get
        Set(value As ViewItemModel)
            Me._MainViewContent = value
            If value Is Nothing Then
                Me.MainContent = Nothing
            Else
                Me.MainContent = value.Content
            End If
        End Set
    End Property

    Private _RightContent As Object
    <JsonIgnore>
    Public Property RightContent As Object
        Get
            Return Me._RightContent
        End Get
        Set(value As Object)
            Me._RightContent = value
            Call _OptimizeDynamicView()
            RaisePropertyChanged("RightContent")
        End Set
    End Property

    Private _RightViewContent As ViewItemModel
    Public Property RightViewContent As ViewItemModel
        Get
            Return Me._RightViewContent
        End Get
        Set(value As ViewItemModel)
            Me._RightViewContent = value
            If value Is Nothing Then
                Me.RightContent = Nothing
            Else
                Me.RightContent = value.Content
            End If
        End Set
    End Property

    Private _BottomContent As Object
    <JsonIgnore>
    Public Property BottomContent As Object
        Get
            Return Me._BottomContent
        End Get
        Set(value As Object)
            Me._BottomContent = value
            Call _OptimizeDynamicView()
            RaisePropertyChanged("BottomContent")
        End Set
    End Property

    Private _BottomViewContent As ViewItemModel
    Public Property BottomViewContent As ViewItemModel
        Get
            Return Me._BottomViewContent
        End Get
        Set(value As ViewItemModel)
            Me._BottomViewContent = value
            If value Is Nothing Then
                Me.BottomContent = Nothing
            Else
                Me.BottomContent = value.Content
            End If
        End Set
    End Property

    Private _IsHorizontalSplitterEnabled As Boolean
    <JsonIgnore>
    Public Property IsHorizontalSplitterEnabled As Boolean
        Get
            Return Me._IsHorizontalSplitterEnabled
        End Get
        Set(value As Boolean)
            Me._IsHorizontalSplitterEnabled = value
            RaisePropertyChanged("IsHorizontalSplitterEnabled")
        End Set
    End Property

    Private _IsVerticalSplitterEnabled As Boolean
    <JsonIgnore>
    Public Property IsVerticalSplitterEnabled As Boolean
        Get
            Return Me._IsVerticalSplitterEnabled
        End Get
        Set(value As Boolean)
            Me._IsVerticalSplitterEnabled = value
            RaisePropertyChanged("IsVerticalSplitterEnabled")
        End Set
    End Property

    Private Sub _OptimizeDynamicView()
        If Me.RightContent Is Nothing Then
            Me.RightGridWidth = New GridLength(0.0)
            Me.HorizontalSplitterWidth = New GridLength(0.0)
            Me.IsHorizontalSplitterEnabled = False
        Else
            If Me.ContentViewWidth > 0.0 Then
                Me.ContentGridWidth = New GridLength(ContentViewWidth)
                Me.RightGridWidth = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
            Else
                Me.RightGridWidth = GridLength.Auto
            End If
            Me.HorizontalSplitterWidth = New GridLength(5.0)
            Me.IsHorizontalSplitterEnabled = True
        End If
        If Me.BottomContent Is Nothing Then
            Me.BottomGridHeight = New GridLength(0.0)
            Me.VerticalSplitterHeight = New GridLength(0.0)
            Me.IsVerticalSplitterEnabled = False
        Else
            If Me.ContentViewHeight > 0.0 Then
                Me.ContentGridHeight = New GridLength(ContentViewHeight)
                Me.BottomGridHeight = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
            Else
                Me.BottomGridHeight = GridLength.Auto
            End If
            Me.VerticalSplitterHeight = New GridLength(5.0)
            Me.IsVerticalSplitterEnabled = True
        End If
    End Sub

    Public Sub New()
        Me.ContentGridHeight = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
        Me.ContentGridWidth = New GridLength(GridLength.Auto.Value, GridUnitType.Star)
    End Sub
End Class
