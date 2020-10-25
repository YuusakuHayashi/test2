'Imports System.ComponentModel

Public Class HistoryViewModel
    Inherits BaseViewModel2

    Private _History As HistoryModel
    Public Property History As HistoryModel
        Get
            Return Me._History
        End Get
        Set(value As HistoryModel)
            Me._History = value
        End Set
    End Property

    'Public Overrides ReadOnly Property FrameType As String
    '    Get
    '        Return ViewModel.HISTORY_FRAME
    '    End Get
    'End Property
    'Inherits BaseViewModel2(Of Object)
    '    Inherits ProjectBaseViewModel(Of HistoryViewModel)

    '    Private _HistoryContents As String
    '    Public Property HistoryContents As String
    '        Get
    '            Return Me._HistoryContents
    '        End Get
    '        Set(value As String)
    '            Me._HistoryContents = value
    '            RaisePropertyChanged("HistoryContents")
    '        End Set
    '    End Property

    '    Private Sub _HistoryContentsUpdate(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
    '        Dim hist As HistoryModel
    '        If e.PropertyName = "Contents" Then
    '            hist = CType(sender, HistoryModel)
    '            Me.HistoryContents = hist.Contents
    '        End If
    '    End Sub

    '    Private Sub _AddHandlerHistoryChanged()
    '        AddHandler Model.History.PropertyChanged, AddressOf _HistoryContentsUpdate
    '    End Sub

    '    Protected Overrides Sub ContextModelCheck()
    '        Dim b As Boolean : b = True
    '        If b Then
    '            ViewModel.SetContext(ViewModel.HISTORY_VIEW, Me.GetType.Name, Me)
    '        End If
    '    End Sub

    '    Protected Overrides Sub MyInitializing()
    '        Me.HistoryContents = Model.History.Contents
    '    End Sub


    '    ' コマンドプロパティ（クリア）
    '    Private _ClearCommand As ICommand
    '    Public ReadOnly Property ClearCommand As ICommand
    '        Get
    '            If Me._ClearCommand Is Nothing Then
    '                Me._ClearCommand = New DelegateCommand With {
    '                    .ExecuteHandler = AddressOf _ClearCommandExecute,
    '                    .CanExecuteHandler = AddressOf _ClearCommandCanExecute
    '                }
    '                Return Me._ClearCommand
    '            Else
    '                Return Me._ClearCommand
    '            End If
    '        End Get
    '    End Property


    '    'コマンド実行可否フラグチェック（クリア）
    '    Private Sub _CheckClearCommandEnabled()
    '        Dim b As Boolean : b = True
    '        Me._ClearEnableFlag = b
    '    End Sub


    '    'コマンド実行可否のフラグ（クリア）
    '    Private __ClearEnableFlag As Boolean
    '    Public Property _ClearEnableFlag As Boolean
    '        Get
    '            Return Me.__ClearEnableFlag
    '        End Get
    '        Set(value As Boolean)
    '            Me.__ClearEnableFlag = value
    '            RaisePropertyChanged("_ClearEnableFlag")
    '            CType(ClearCommand, DelegateCommand).RaiseCanExecuteChanged()
    '        End Set
    '    End Property


    '    ' コマンド実行（クリア）
    '    Private Sub _ClearCommandExecute(ByVal parameter As Object)
    '        Model.History.ClearContents()
    '    End Sub


    '    ' コマンド実行可否（クリア）
    '    Private Function _ClearCommandCanExecute(ByVal parameter As Object) As Boolean
    '        Return Me._ClearEnableFlag
    '    End Function


    '    Sub New(ByRef m As Model,
    '            ByRef vm As ViewModel)

    '        Dim ccep(0) As CheckCommandEnabledProxy
    '        Dim ccep2 As CheckCommandEnabledProxy

    '        Dim ahp(0) As AddHandlerProxy
    '        Dim ahp2 As AddHandlerProxy

    '        ccep(0) = AddressOf Me._CheckClearCommandEnabled
    '        ccep2 = [Delegate].Combine(ccep)

    '        ahp(0) = AddressOf Me._AddHandlerHistoryChanged
    '        ahp2 = [Delegate].Combine(ahp)

    '        ' ビューモデルの設定
    '        Initializing(m, vm, ccep2, ahp2)
    '    End Sub
    Private Sub _ViewInitialize()
        Me.History = AppInfo.ProjectInfo.Model.Data.History
    End Sub

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel,
                                    ByRef vm As ViewModel)
        InitializeHandler = AddressOf _ViewInitialize

        Call BaseInitialize(app, vm)
    End Sub
End Class
