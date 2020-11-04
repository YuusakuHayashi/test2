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

    '---------------------------------------------------------------------------------------------'
    ' 新仕様
    Private _DynamicView As DynamicViewModel
    Public Property DynamicView As DynamicViewModel
        Get
            If Me._DynamicView Is Nothing Then
                Me._DynamicView = New DynamicViewModel
            End If
            Return Me._DynamicView
        End Get
        Set(value As DynamicViewModel)
            Me._DynamicView = value
        End Set
    End Property
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    ' 旧仕様
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
    '---------------------------------------------------------------------------------------------'

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

    Private _FontSize As Double
    Public Property [FontSize] As Double
        Get
            If Me._FontSize = 0.0 Then
                Me._FontSize = 11.0
            End If
            Return Me._FontSize
        End Get
        Set(value As Double)
            Me._FontSize = value
            RaisePropertyChanged("FontSize")
        End Set
    End Property

    'Private _ViewExplorerStretchType As String
    'Public Property ViewExplorerStretchType As String
    '    Get
    '        Return _ViewExplorerStretchType
    '    End Get
    '    Set(value As String)
    '        _ViewExplorerStretchType = value
    '        Call _SwitchViewExplorerStretch()
    '        RaisePropertyChanged("ViewExplorerStretchType")
    '    End Set
    'End Property

    'Private _ViewExplorerStretch As Stretch
    'Public Property ViewExplorerStretch As Stretch
    '    Get
    '        Return _ViewExplorerStretch
    '    End Get
    '    Set(value As Stretch)
    '        _ViewExplorerStretch = value
    '    End Set
    'End Property

    'Private Sub _SwitchViewExplorerStretch()
    '    Select Case ViewExplorerStretchType
    '        Case "None"
    '            ViewExplorerStretch = Stretch.None
    '        Case "Fill"
    '            ViewExplorerStretch = Stretch.Fill
    '        Case "Uniform"
    '            ViewExplorerStretch = Stretch.Uniform
    '        Case "UniformToFill"
    '            ViewExplorerStretch = Stretch.UniformToFill
    '    End Select
    'End Sub

    Public Const SINGLE_VIEW As String = "Single"
    Public Const MULTI_VIEW As String = "Multi"
End Class
