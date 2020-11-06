Imports test2
Imports System.Collections.ObjectModel

Public Class RpaProjectModel : Inherits ProjectModel

    Public RootDirectoryName As String
    Public UserDirectoryName As String
    Public SystemDirectoryName As String
    Public PythonPathName As String

    Private _Rpas As ObservableCollection(Of RpaModel)
    Public Property Rpas As ObservableCollection(Of RpaModel)
        Get
            If Me._Rpas Is Nothing Then
                Me._Rpas = New ObservableCollection(Of RpaModel)
            End If
            Return Me._Rpas
        End Get
        Set(value As ObservableCollection(Of RpaModel))
            Me._Rpas = value
        End Set
    End Property
    Public ReadOnly Property RootSysDirectoryName As String
        Get
            Return Me.RootDirectoryName & "\sys"
        End Get
    End Property
    Public ReadOnly Property SysDirectoryName As String
        Get
            Return Me.SystemDirectoryName & "\sys"
        End Get
    End Property
    Public ReadOnly Property MyDirFileName As String
        Get
            Return Me.SystemDirectoryName & "\mydir"
        End Get
    End Property
    Public ReadOnly Property MyPythonFileName As String
        Get
            Return Me.SystemDirectoryName & "\mypython"
        End Get
    End Property
    Public ReadOnly Property RpaProjectFileName As String
        Get
            Return Me.SystemDirectoryName & "\rpa_project.yaml"
        End Get
    End Property

    Public Sub ProjectAdd(ByVal [name] As String)
    End Sub

    Public Function GetRpaIndex() As Integer
        Dim i = 0
        Dim b = False
        Do Until True = False
            b = False
            For Each rpa In Me.Rpas
                If i = rpa.Index Then
                    i += 1
                    b = True
                    Exit For
                End If
            Next
            If Not b Then
                Exit Do
            End If
        Loop
        GetRpaIndex = i
    End Function

    Public Overrides Sub ViewSetupExecute(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        Dim mvm = New MenuViewModel
        Dim vevm = New ViewExplorerViewModel
        Dim rpapvm = New RpaProjectViewModel
        Dim rpapmvm = New RpaProjectMenuViewModel
        Call mvm.Initialize(app, vm)
        Call vevm.Initialize(app, vm)
        Call rpapvm.Initialize(app, vm)
        Call rpapmvm.Initialize(app, vm)
        Dim dvm = New FlexibleViewModel With {
            .ContentViewHeight = 25.0,
            .MainViewContent = New ViewItemModel With {
                .Name = "メニュー",
                .Content = mvm
            },
            .BottomViewContent = New ViewItemModel With {
                .Name = vbNullString,
                .Content = New FlexibleViewModel With {
                    .ContentViewWidth = 200.0,
                    .MainViewContent = New ViewItemModel With {
                        .Name = "ビューエクスプローラー",
                        .Content = vevm
                    },
                    .RightViewContent = New ViewItemModel With {
                        .Name = vbNullString,
                        .Content = New FlexibleViewModel With {
                            .ContentViewHeight = 25.0,
                            .MainViewContent = New ViewItemModel With {
                                .Name = "Rpaプロジェクトメニュー",
                                .Content = rpapmvm
                            },
                            .BottomViewContent = New ViewItemModel With {
                                .Name = "Rpaプロジェクト",
                                .Content = rpapvm
                            }
                        }
                    }
                }
            }
        }
        vm.VisualizeView(dvm)
    End Sub

    Public Overrides Function ViewDefineExecute(ByVal mname As String) As Object
        Dim obj As Object
        Select Case mname
            Case (New AppDirectoryModel).GetType.Name
                obj = New AppDirectoryModel
            Case (New ViewModel).GetType.Name
                obj = New ViewModel
            Case (New HistoryViewModel).GetType.Name
                obj = New HistoryViewModel
            Case (New MenuViewModel).GetType.Name
                obj = New MenuViewModel
            Case (New RpaProjectMenuViewModel).GetType.Name
                obj = New RpaProjectMenuViewModel
            Case (New RpaProjectViewModel).GetType.Name
                obj = New RpaProjectViewModel
            Case (New RpaViewModel).GetType.Name
                obj = New RpaViewModel
            Case Else
                obj = Nothing
        End Select

        ViewDefineExecute = obj
    End Function

    ''--- ビュー削除 ------------------------------------------------------------------------------'
    ' 廃止
    '' ビューに削除に紐づいたモデルを削除がする
    'Private Sub _DeleteViewAddHandler()
    '    AddHandler _
    '        DelegateEventListener.Instance.DeleteViewRequested,
    '        AddressOf Me._DeleteViewRequestedReview
    'End Sub

    'Private Sub _DeleteViewRequestedReview(ByVal [view] As ViewItemModel, ByVal e As System.EventArgs)
    '    Call _DeleteViewRequestAccept([view])
    'End Sub

    'Private Sub _DeleteViewRequestAccept(ByVal [view] As ViewItemModel)
    '    If [view].ModelName = (New RpaViewModel).GetType.Name Then
    '        For Each rpa In Me.Rpas
    '            If rpa.Equals([view].Content) Then
    '                Me.Rpas.Remove(rpa)
    '                Exit For
    '            End If
    '        Next
    '    End If
    'End Sub
    '---------------------------------------------------------------------------------------------'
End Class
