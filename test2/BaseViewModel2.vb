Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class BaseViewModel2(Of T As {New})
    Implements INotifyPropertyChanged
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

    Private _Data As T
    Public Property Data As T
        Get
            Return _Data
        End Get
        Set(value As T)
            _Data = value
        End Set
    End Property

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

    Protected Delegate Sub InitializeProxy()
    Protected Delegate Sub CheckCommandEnabledProxy()

    Protected Overridable Sub ViewInitializing()
        'Nothing To Do
    End Sub

    Protected Overridable Sub ContextModelCheck()
        'Nothing To Do
    End Sub


    ' ビューモデル初期化時のメインメソッド
    ' 共有のモデル・ビューモデル・アプリケーション情報、プロジェクト情報をセットする
    Protected Overloads Sub Initializing(ByVal m As Model,
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

        If Me.Model.Data <> Nothing Then
            obj = Me.Model.Data
            Select Case obj.GetType
                Case (New Object).GetType
                    Me.Data = CType(obj, T)
                Case (New JObject).GetType
                    ' Ｊｓｏｎからロードした場合は、JObject型になっている
                    Me.Data = obj.ToObject(Of T)
                Case (New T).GetType
                    Me.Data = obj
            End Select
        Else
            Me.Data = New T
        End If

        If ih <> Nothing Then
            Call ih()
        End If
        If cceh <> Nothing Then
            Call cceh()
        End If
        If ah <> Nothing Then
            Call ah()
        End If
    End Sub

    'Protected Overloads Sub Initializing(ByVal m As Model,
    '                                     ByVal vm As ViewModel,
    '                                     ByVal adm As AppDirectoryModel,
    '                                     ByVal pim As ProjectInfoModel,
    '                                     ByRef ip As InitializeProxy)
    '    Call Initializing(m, vm, adm, pim)

    '    ' 自身(ビューモデル）の初期化設定を行います
    '    Call ip()
    'End Sub

    'Protected Overloads Sub Initializing(ByVal m As Model,
    '                                     ByVal vm As ViewModel,
    '                                     ByVal adm As AppDirectoryModel,
    '                                     ByVal pim As ProjectInfoModel,
    '                                     ByRef ip As InitializeProxy,
    '                                     ByRef ccep As CheckCommandEnabledProxy)
    '    Call Initializing(m, vm, adm, pim, ip)

    '    ' コマンド実行可否の設定
    '    Call ccep()
    'End Sub

    'Protected Overloads Sub Initializing(ByVal m As Model,
    '                                     ByVal vm As ViewModel,
    '                                     ByVal adm As AppDirectoryModel,
    '                                     ByVal pim As ProjectInfoModel,
    '                                     ByRef ip As InitializeProxy,
    '                                     ByRef ahp As AddHandlerProxy)
    '    Call Initializing(m, vm, adm, pim, ip)

    '    ' イベントハンドラの登録
    '    Call ahp()
    'End Sub

    'Protected Overloads Sub Initializing(ByVal m As Model,
    '                                     ByVal vm As ViewModel,
    '                                     ByVal adm As AppDirectoryModel,
    '                                     ByVal pim As ProjectInfoModel,
    '                                     ByRef ip As InitializeProxy,
    '                                     ByRef ccep As CheckCommandEnabledProxy,
    '                                     ByRef ahp As AddHandlerProxy)
    '    Call Initializing(m, vm, adm, pim, ip, ccep)

    '    ' イベントハンドラの登録
    '    Call ahp()
    'End Sub
End Class
