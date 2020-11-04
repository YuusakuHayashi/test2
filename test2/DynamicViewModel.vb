Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class DynamicViewModel
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Public Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub

    'Private _TopView As DynamicViewModel
    'Public Property TopView As DynamicViewModel
    '    Get
    '        If Me._TopView Is Nothing Then
    '            Me._TopView = New DynamicViewModel
    '        End If
    '        Return Me._TopView
    '    End Get
    '    Set(value As DynamicViewModel)
    '        Me._TopView = value
    '        RaisePropertyChanged("TopView")
    '    End Set
    'End Property

    'Private _ButtomView As DynamicViewModel
    'Public Property ButtomView As DynamicViewModel
    '    Get
    '        If Me._ButtomView Is Nothing Then
    '            Me._ButtomView = New DynamicViewModel
    '        End If
    '        Return Me._ButtomView
    '    End Get
    '    Set(value As DynamicViewModel)
    '        Me._ButtomView = value
    '        RaisePropertyChanged("ButtomView")
    '    End Set
    'End Property

    'Private _LeftView As DynamicViewModel
    'Public Property LeftView As DynamicViewModel
    '    Get
    '        If Me._LeftView Is Nothing Then
    '            Me._LeftView = New DynamicViewModel
    '        End If
    '        Return Me._LeftView
    '    End Get
    '    Set(value As DynamicViewModel)
    '        Me._LeftView = value
    '        RaisePropertyChanged("LeftView")
    '    End Set
    'End Property

    'Private _RightView As DynamicViewModel
    'Public Property RightView As DynamicViewModel
    '    Get
    '        If Me._RightView Is Nothing Then
    '            Me._RightView = New DynamicViewModel
    '        End If
    '        Return Me._RightView
    '    End Get
    '    Set(value As DynamicViewModel)
    '        Me._RightView = value
    '        RaisePropertyChanged("RightView")
    '    End Set
    'End Property

    Private _ButtomViewHeight As Double
    Public Property ButtomViewHeight As Double
        Get
            Return Me._ButtomViewHeight
        End Get
        Set(value As Double)
            Me._ButtomViewHeight = value
        End Set
    End Property

    Private _ButtomViewPreservedHeight As Double
    Public Property ButtomViewPreservedHeight As Double
        Get
            Return Me._ButtomViewPreservedHeight
        End Get
        Set(value As Double)
            Me._ButtomViewPreservedHeight = value
        End Set
    End Property

    Private _ButtomGridHeight As GridLength
    <JsonIgnore>
    Public Property ButtomGridHeight As GridLength
        Get
            Return Me._ButtomGridHeight
        End Get
        Set(value As GridLength)
            Me._ButtomGridHeight = value
            Me.ButtomViewHeight = value.Value
            RaisePropertyChanged("ButtomGridHeight")
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

    Private _TopLeftContent As Object
    <JsonIgnore>
    Public Property TopLeftContent As Object
        Get
            Return Me._TopLeftContent
        End Get
        Set(value As Object)
            Me._TopLeftContent = value
            Call _OptimizeDynamicView()
            RaisePropertyChanged("TopLeftContent")
        End Set
    End Property

    Private _TopLeftViewContent As Object
    Public Property TopLeftViewContent As Object
        Get
            Return Me._TopLeftViewContent
        End Get
        Set(value As Object)
            Me._TopLeftViewContent = value
            Me.TopLeftContent = _TakeOutContent(value)
        End Set
    End Property

    Private _TopRightContent As Object
    <JsonIgnore>
    Public Property TopRightContent As Object
        Get
            Return Me._TopRightContent
        End Get
        Set(value As Object)
            Me._TopRightContent = value
            Call _OptimizeDynamicView()
            RaisePropertyChanged("TopRightContent")
        End Set
    End Property

    Private _TopRightViewContent As Object
    Public Property TopRightViewContent As Object
        Get
            Return Me._TopRightViewContent
        End Get
        Set(value As Object)
            Me._TopRightViewContent = value
            Me.TopRightContent = _TakeOutContent(value)
        End Set
    End Property

    Private _ButtomLeftContent As Object
    <JsonIgnore>
    Public Property ButtomLeftContent As Object
        Get
            Return Me._ButtomLeftContent
        End Get
        Set(value As Object)
            Me._ButtomLeftContent = value
            Call _OptimizeDynamicView()
            RaisePropertyChanged("ButtomLeftContent")
        End Set
    End Property

    Private _ButtomLeftViewContent As Object
    Public Property ButtomLeftViewContent As Object
        Get
            Return Me._ButtomLeftViewContent
        End Get
        Set(value As Object)
            Me._ButtomLeftViewContent = value
            Me.ButtomLeftContent = _TakeOutContent(value)
        End Set
    End Property

    Private Function _TakeOutContent(value)
        If TryCast(value, ViewItemModel) IsNot Nothing Then
            Return value.Content
        Else
            Throw New Exception("DynamicViewModel TakeOutContent!!")
        End If
    End Function

    Private Sub _OptimizeDynamicView()
        If Me.TopRightViewContent Is Nothing Then
            Me.RightGridWidth = New GridLength(0.0)
        Else
            Me.RightGridWidth = GridLength.Auto
        End If
        If Me.ButtomLeftViewContent Is Nothing Then
            Me.ButtomGridHeight = New GridLength(0.0)
        Else
            Me.ButtomGridHeight = GridLength.Auto
        End If
    End Sub
End Class
