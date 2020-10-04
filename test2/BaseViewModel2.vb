Imports System.ComponentModel

Public Class BaseViewModel2(Of T)
    Inherits ModelLoader(Of T)
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

    'Private _Model As Object
    'Public Property Model As Object
    '    Get
    '        Return _Model
    '    End Get
    '    Set(value As Object)
    '        _Model = value
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

    Protected Delegate Sub InitializingProxy()
    Protected Delegate Sub ContextModelCheckProxy()
    Protected Delegate Sub ViewModelInitializingProxy()
    Protected Delegate Sub CheckCommandEnabledProxy()
    Protected Delegate Sub AddHandlerProxy()

    Protected Overridable Sub ViewInitializing()
        'Nothing To Do
    End Sub

    Protected Overridable Sub ContextModelCheck()
        'Nothing To Do
    End Sub


    Protected Overloads Sub Initializing(ByRef m As Model,
                                         ByRef vm As ViewModel,
                                         ByRef adm As AppDirectoryModel,
                                         ByRef pim As ProjectInfoModel)
        Model = m
        ViewModel = vm
        AppInfo = adm
        ProjectInfo = pim
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model,
                                         ByRef vm As ViewModel,
                                         ByRef adm As AppDirectoryModel,
                                         ByRef pim As ProjectInfoModel,
                                         ByRef cmcp As ContextModelCheckProxy)
        Initializing(m, vm, adm, pim)

        ' ビューモデルの設定（画面登録など）を行います
        cmcp()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model,
                                         ByRef vm As ViewModel,
                                         ByRef adm As AppDirectoryModel,
                                         ByRef pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef cmcp As ContextModelCheckProxy)
        Initializing(m, vm, adm, pim)

        ' 自身(ビューモデル）の初期化設定を行います
        ip()

        ' ビューモデルの設定（画面登録など）を行います
        cmcp()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model,
                                         ByRef vm As ViewModel,
                                         ByRef adm As AppDirectoryModel,
                                         ByRef pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef cmcp As ContextModelCheckProxy,
                                         ByRef cccep As CheckCommandEnabledProxy)
        Initializing(m, vm, adm, pim, ip, cmcp)

        ' コマンド実行可否の設定
        cccep()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model,
                                         ByRef vm As ViewModel,
                                         ByRef adm As AppDirectoryModel,
                                         ByRef pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef cmcp As ContextModelCheckProxy,
                                         ByRef ahp As AddHandlerProxy)
        Initializing(m, vm, adm, pim, ip, cmcp)

        ' イベントハンドラの登録
        ahp()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model,
                                         ByRef vm As ViewModel,
                                         ByRef adm As AppDirectoryModel,
                                         ByRef pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef cmcp As ContextModelCheckProxy,
                                         ByRef cccep As CheckCommandEnabledProxy,
                                         ByRef ahp As AddHandlerProxy)
        Initializing(m, vm, adm, pim, ip, cmcp, cccep)

        ' イベントハンドラの登録
        ahp()
    End Sub
End Class
