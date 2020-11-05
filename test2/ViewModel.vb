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

    Public Sub ShowDynamicView(ByVal dvm As DynamicViewModel)
        Me.Views = Nothing
        Call _CreateViews(dvm)
        Me.Content = dvm
    End Sub

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

    Private Overloads Sub _CreateViews(ByVal dvm As DynamicViewModel)
        'Dim act As Action(Of Object, String)
        'act = Sub(ByVal obj As Object, ByVal [name] As String)
        '          Dim dvm2 As DynamicViewModel
        '          Dim tvm As TabViewModel
        '          Select Case obj.GetType.Name
        '              Case "DynamicViewModel"
        '                  dvm2 = CType(obj, DynamicViewModel)
        '                  Call _CreateViews(dvm2)
        '              Case "TabViewModel"
        '                  tvm = CType(obj, TabViewModel)
        '                  Call _CreateViews(tvm)
        '              Case Else
        '                  ViewModel.Views.Add(
        '                     New ViewItemModel With {
        '                         .Name = [name]
        '                     }
        '                 )
        '          End Select
        '      End Sub
        Dim act As Action(Of ViewItemModel)
        act = Sub(ByVal v As ViewItemModel)
                  Dim dvm2 As DynamicViewModel
                  Dim tvm As TabViewModel
                  Select Case v.Content.GetType.Name
                      Case "DynamicViewModel"
                          dvm2 = CType(v.Content, DynamicViewModel)
                          Call _CreateViews(dvm2)
                      Case "TabViewModel"
                          tvm = CType(v.Content, TabViewModel)
                          Call _CreateViews(tvm)
                      Case Else
                          Me.Views.Add(v)
                  End Select
              End Sub
        Call act(dvm.MainViewContent)
        If dvm.RightContent IsNot Nothing Then
            Call act(dvm.RightViewContent)
        End If
        If dvm.BottomContent IsNot Nothing Then
            Call act(dvm.BottomViewContent)
        End If
    End Sub

    Private Overloads Sub _CreateViews(ByVal tvm As TabViewModel)
        Dim dvm As DynamicViewModel
        Dim tvm2 As TabViewModel
        For Each t In tvm.Tabs
            Select Case t.Content.GetType.Name
                Case "DynamicViewModel"
                    dvm = CType(t.Content, DynamicViewModel)
                    Call _CreateViews(dvm)
                Case "TabViewModel"
                    tvm2 = CType(t.Content, TabViewModel)
                    Call _CreateViews(tvm2)
                Case Else
                    Me.Views.Add(
                        New ViewItemModel With {
                            .Name = t.Name
                        }
                    )
            End Select
        Next
    End Sub

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
