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

    Private _WindowHeight As Double
    Public Property WindowHeight As Double
        Get
            If Me._WindowHeight = 0.0 Then
                Me._WindowHeight = 500.0
            End If
            Return Me._WindowHeight
        End Get
        Set(value As Double)
            Me._WindowHeight = value
            RaisePropertyChanged("WindowHeight")
        End Set
    End Property

    Private _WindowWidth As Double
    Public Property WindowWidth As Double
        Get
            If Me._WindowWidth = 0.0 Then
                Me._WindowWidth = 1000.0
            End If
            Return Me._WindowWidth
        End Get
        Set(value As Double)
            Me._WindowWidth = value
            RaisePropertyChanged("WindowWidth")
        End Set
    End Property

    Private _Content As FrameViewModel
    Public Property Content As FrameViewModel
        Get
            Return Me._Content
        End Get
        Set(value As FrameViewModel)
            Me._Content = value
            RaisePropertyChanged("Content")
        End Set
    End Property

    Private _SaveContent As FrameViewModel
    <JsonIgnore>
    Public Property SaveContent As FrameViewModel
        Get
            Return Me._SaveContent
        End Get
        Set(value As FrameViewModel)
            Me._SaveContent = value
        End Set
    End Property

    'Private _Content As FlexibleViewModel
    'Public Property Content As FlexibleViewModel
    '    Get
    '        Return Me._Content
    '    End Get
    '    Set(value As FlexibleViewModel)
    '        Me._Content = value
    '        RaisePropertyChanged("Content")
    '    End Set
    'End Property

    'Private _SaveContent As FlexibleViewModel
    '<JsonIgnore>
    'Public Property SaveContent As FlexibleViewModel
    '    Get
    '        Return Me._SaveContent
    '    End Get
    '    Set(value As FlexibleViewModel)
    '        Me._SaveContent = value
    '    End Set
    'End Property

    'Public Sub VisualizeView(ByVal fvm As FlexibleViewModel)
    '    Me.Views = Nothing
    '    Me.Views = _CreateViews(fvm, Me.Views)
    '    Me.Content = fvm
    'End Sub
    Public Sub VisualizeView(ByVal fvm As FrameViewModel)
        Me.Views = Nothing
        Me.Content = fvm
    End Sub

    '---------------------------------------------------------------------------------------------'
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
            Call DelegateEventListener.Instance.RaiseViewsChanged()
        End Set
    End Property
    '---------------------------------------------------------------------------------------------'

    Public Sub ReloadViews()
        'Me.Views = Nothing
        'Me.Views = _CreateViews(Me.Content, Me.Views)
    End Sub

    Private Overloads Function _CreateViews(ByVal fvm As FlexibleViewModel,
                                            ByVal vims As ObservableCollection(Of ViewItemModel)) As ObservableCollection(Of ViewItemModel)

        'Dim fnc As Func(Of ViewItemModel, ObservableCollection(Of ViewItemModel), ObservableCollection(Of ViewItemModel))
        'fnc = Function(ByVal v As ViewItemModel, ByVal vs As ObservableCollection(Of ViewItemModel)) As ObservableCollection(Of ViewItemModel)
        '          Dim fvm2 As FlexibleViewModel
        '          Dim tvm As TabViewModel
        '          Select Case v.Content.GetType.Name
        '              Case "FlexibleViewModel"
        '                  fvm2 = CType(v.Content, FlexibleViewModel)
        '                  vs = _CreateViews(fvm2, vs)
        '              Case "TabViewModel"
        '                  tvm = CType(v.Content, TabViewModel)
        '                  vs = _CreateViews(tvm, vs)
        '              Case Else
        '                  vs.Add(v)
        '          End Select
        '          Return vs
        '      End Function

        'If fvm.MainContent IsNot Nothing Then
        '    vims = fnc(fvm.MainViewContent, vims)
        'End If
        'If fvm.RightContent IsNot Nothing Then
        '    vims = fnc(fvm.RightViewContent, vims)
        'End If
        'If fvm.BottomContent IsNot Nothing Then
        '    vims = fnc(fvm.BottomViewContent, vims)
        'End If
        '_CreateViews = vims
    End Function

    Private Overloads Function _CreateViews(ByVal tvm As TabViewModel,
                                            ByVal vims As ObservableCollection(Of ViewItemModel)) As ObservableCollection(Of ViewItemModel)
        'Dim fvm As FlexibleViewModel
        'Dim tvm2 As TabViewModel
        'For Each vt In tvm.ViewContentTabs
        '    Select Case vt.Content.GetType.Name
        '        Case "FlexibleViewModel"
        '            fvm = CType(vt.Content, FlexibleViewModel)
        '            vims = _CreateViews(fvm, vims)
        '        Case "TabViewModel"
        '            tvm2 = CType(vt.Content, TabViewModel)
        '            vims = _CreateViews(tvm2, vims)
        '        Case Else
        '            vims.Add(vt)
        '    End Select
        'Next
        '_CreateViews = vims
    End Function

    'Public Const SINGLE_VIEW As String = "Single"
    'Public Const MULTI_VIEW As String = "Multi"
End Class
