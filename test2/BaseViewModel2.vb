Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

'Public Class BaseViewModel2(Of T As {New})
Public MustInherit Class BaseViewModel2
    Implements INotifyPropertyChanged, BaseViewModelInterface

    '--- INortify -------------------------------------------------------------------------------------'
    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub
    '--------------------------------------------------------------------------------------------------'
    Private _Model As Model
    Public Property Model As Model
        Get
            Return _Model
        End Get
        Set(value As Model)
            _Model = value
        End Set
    End Property

    'Private _Data As T
    'Public Property Data As T
    '    Get
    '        Return _Data
    '    End Get
    '    Set(value As T)
    '        _Data = value
    '        Model.Data = Data
    '    End Set
    'End Property

    Private _ViewModel As ViewModel
    Public Property ViewModel As ViewModel
        Get
            Return _ViewModel
        End Get
        Set(value As ViewModel)
            _ViewModel = value
        End Set
    End Property

    Private _AppInfo As AppDirectoryModel
    Public Property AppInfo As AppDirectoryModel
        Get
            Return _AppInfo
        End Get
        Set(value As AppDirectoryModel)
            _AppInfo = value
        End Set
    End Property

    Private _ProjectInfo As ProjectInfoModel
    Public Property ProjectInfo As ProjectInfoModel
        Get
            Return _ProjectInfo
        End Get
        Set(value As ProjectInfoModel)
            _ProjectInfo = value
        End Set
    End Property

    Protected Property InitializeHandler As Action
    Protected Property CheckCommandEnabledHandler As Action
    Protected Property [AddHandler] As Action

    Public MustOverride ReadOnly Property FrameType As String Implements BaseViewModelInterface.FrameType

    Protected Overridable Sub ViewInitializing()
        'Nothing To Do
    End Sub

    Protected Overridable Sub ContextModelCheck()
        'Nothing To Do
    End Sub


    ' Model型のDataObjectはJsonからデシリアライズされたものはJObject型になっている
    Protected Sub BaseInitialize(ByVal m As Model,
                                 ByVal vm As ViewModel,
                                 ByVal adm As AppDirectoryModel,
                                 ByVal pim As ProjectInfoModel)
        Dim obj As Object
        Dim ih = Me.InitializeHandler
        Dim cceh = Me.CheckCommandEnabledHandler
        Dim ah = Me.[AddHandler]

        Me.Model = m
        Me.ViewModel = vm
        Me.AppInfo = adm
        Me.ProjectInfo = pim

        'If Me.Model.Data IsNot Nothing Then
        '    obj = Me.Model.Data
        '    Select Case obj.GetType
        '        Case (New Object).GetType
        '            Me.Model.Data = CType(obj, T)
        '        Case (New JObject).GetType
        '            ' Ｊｓｏｎからロードした場合は、JObject型になっている
        '            Me.Model.Data = obj.ToObject(Of T)
        '        Case (New T).GetType
        '            Me.Model.Data = obj
        '    End Select
        'Else
        '    Me.Model.Data = New T
        'End If

        ' 自身(ビューモデル）の初期化設定を行います
        If ih <> Nothing Then
            Call ih()
        End If
        ' コマンド実行可否の設定
        If cceh <> Nothing Then
            Call cceh()
        End If
        ' イベントハンドラの登録
        If ah <> Nothing Then
            Call ah()
        End If
    End Sub

    Public MustOverride Sub Initialize(ByRef m As Model, ByRef vm As ViewModel, ByRef adm As AppDirectoryModel, ByRef pim As ProjectInfoModel) Implements BaseViewModelInterface.Initialize
End Class
