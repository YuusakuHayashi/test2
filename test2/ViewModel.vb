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
    Public Property Content As Object
        Get
            Return Me._Content
        End Get
        Set(value As Object)
            Me._Content = value
            RaisePropertyChanged("Content")
        End Set
    End Property

    'Private _FlexibleStructure As FlexibleStructureModel
    'Public Property FlexibleStructure As FlexibleStructureModel
    '    Get
    '        Return Me._FlexibleStructure
    '    End Get
    '    Set(value As FlexibleStructureModel)
    '        Me._FlexibleStructure = value
    '    End Set
    'End Property

    Private _SaveContent As Object
    <JsonIgnore>
    Public Property SaveContent As Object
        Get
            Return Me._SaveContent
        End Get
        Set(value As Object)
            Me._SaveContent = value
        End Set
    End Property

    Public Sub VisualizeView(ByVal fvm As FlexibleViewModel)
        Me.Views = Nothing
        Call _CreateViews(fvm)
        Me.Content = fvm
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

    Private Overloads Sub _CreateViews(ByVal fvm As FlexibleViewModel)
        Dim act As Action(Of ViewItemModel)
        act = Sub(ByVal v As ViewItemModel)
                  Dim fvm2 As FlexibleViewModel
                  Dim tvm As TabViewModel
                  Select Case v.Content.GetType.Name
                      Case "FlexibleViewModel"
                          fvm2 = CType(v.Content, FlexibleViewModel)
                          Call _CreateViews(fvm2)
                      Case "TabViewModel"
                          tvm = CType(v.Content, TabViewModel)
                          Call _CreateViews(tvm)
                      Case Else
                          Me.Views.Add(v)
                  End Select
              End Sub
        Call act(fvm.MainViewContent)
        If fvm.RightContent IsNot Nothing Then
            Call act(fvm.RightViewContent)
        End If
        If fvm.BottomContent IsNot Nothing Then
            Call act(fvm.BottomViewContent)
        End If
    End Sub

    Private Overloads Sub _CreateViews(ByVal tvm As TabViewModel)
        Dim fvm As FlexibleViewModel
        Dim tvm2 As TabViewModel
        For Each t In tvm.Tabs
            Select Case t.Content.GetType.Name
                Case "FlexibleViewModel"
                    fvm = CType(t.Content, FlexibleViewModel)
                    Call _CreateViews(fvm)
                Case "TabViewModel"
                    tvm2 = CType(t.Content, TabViewModel)
                    Call _CreateViews(tvm2)
                Case Else
                    Me.Views.Add(
                        New ViewItemModel With {
                            .Name = t.Name,
                            .Content = t.Content
                        }
                    )
            End Select
        Next
    End Sub

    'Private Overloads Sub _CreateFlexibleStructure(ByVal fvm As FlexibleViewModel, ByRef fsm As FlexibleStructureModel)
    '    Dim act As Action(Of ViewItemModel, FlexibleStructureModel)
    '    act = Sub(ByVal vim As ViewItemModel, ByVal fsm2 As FlexibleStructureModel)
    '              Dim fvm2 As FlexibleViewModel
    '              Dim tvm As TabViewModel
    '              Select Case vim.Content.GetType.Name
    '                  Case "FlexibleViewModel"
    '                      fvm2 = CType(vim.Content, FlexibleViewModel)
    '                      Call _CreateFlexibleStructure(fvm2, fsm2)
    '                  Case "TabViewModel"
    '                      tvm = CType(vim.Content, TabViewModel)
    '                      Call _CreateFlexibleStructure(tvm, fsm2)
    '                  Case Else
    '                      fsm2.Name = vim.Name
    '                      fsm2.ModelName = vim.ModelName
    '              End Select
    '          End Sub
    '    Call act(fvm.MainViewContent, fsm.MainStructure)
    '    If fvm.RightContent IsNot Nothing Then
    '        Call act(fvm.RightViewContent, fsm.RightStructure)
    '    End If
    '    If fvm.BottomContent IsNot Nothing Then
    '        Call act(fvm.BottomViewContent, fsm.BottomStructure)
    '    End If

    '    Me.FlexibleStructure = fsm
    'End Sub

    'Private Overloads Sub _CreateFlexibleStructure(ByVal tvm As TabViewModel, ByRef fsm As FlexibleStructureModel)
    '    Dim fvm As FlexibleViewModel
    '    Dim tvm2 As TabViewModel
    '    For Each t In tvm.Tabs
    '        Select Case t.Content.GetType.Name
    '            Case "FlexibleViewModel"
    '                fvm = CType(t.Content, FlexibleViewModel)
    '                Call _CreateFlexibleStructure(fvm, fsm)
    '            Case "TabViewModel"
    '                tvm2 = CType(t.Content, TabViewModel)
    '                Call _CreateFlexibleStructure(tvm2, fsm)
    '            Case Else
    '                fsm.Name = t.Name
    '                fsm.ModelName = t.ModelName
    '        End Select
    '    Next
    'End Sub

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
