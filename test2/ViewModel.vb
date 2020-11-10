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

    'Private _WindowWidth As Double
    'Public Property WindowWidth As Double
    '    Get
    '        If Me._WindowWidth = 0.0 Then
    '            Me._WindowWidth = 1000.0
    '        End If
    '        Return Me._WindowWidth
    '    End Get
    '    Set(value As Double)
    '        Me._WindowWidth = value
    '        RaisePropertyChanged("WindowWidth")
    '    End Set
    'End Property

    'Private _WindowHeight As Double
    'Public Property WindowHeight As Double
    '    Get
    '        If Me._WindowHeight = 0.0 Then
    '            Me._WindowHeight = 1000.0
    '        End If
    '        Return Me._WindowHeight
    '    End Get
    '    Set(value As Double)
    '        Me._WindowHeight = value
    '        RaisePropertyChanged("WindowHeight")
    '    End Set
    'End Property

    Private _Content As FlexibleViewModel
    Public Property Content As FlexibleViewModel
        Get
            Return Me._Content
        End Get
        Set(value As FlexibleViewModel)
            Me._Content = value
            RaisePropertyChanged("Content")
        End Set
    End Property

    Private _SaveContent As FlexibleViewModel
    <JsonIgnore>
    Public Property SaveContent As FlexibleViewModel
        Get
            Return Me._SaveContent
        End Get
        Set(value As FlexibleViewModel)
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
            Call DelegateEventListener.Instance.RaiseViewsChanged()
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
        For Each vt In tvm.ViewContentTabs
            Select Case vt.Content.GetType.Name
                Case "FlexibleViewModel"
                    fvm = CType(vt.Content, FlexibleViewModel)
                    Call _CreateViews(fvm)
                Case "TabViewModel"
                    tvm2 = CType(vt.Content, TabViewModel)
                    Call _CreateViews(tvm2)
                Case Else
                    Me.Views.Add(
                        New ViewItemModel With {
                            .Name = vt.Name,
                            .ModelName = vt.ModelName
                        }
                    )
            End Select
        Next
    End Sub

    Public Const SINGLE_VIEW As String = "Single"
    Public Const MULTI_VIEW As String = "Multi"
End Class
