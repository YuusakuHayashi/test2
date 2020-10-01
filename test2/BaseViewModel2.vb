Imports System.ComponentModel

Public Class BaseViewModel2 : Inherits ProjectModel2
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

    Private _Model As Model2
    Public Property Model As Model2
        Get
            Return _Model
        End Get
        Set(value As Model2)
            _Model = value
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


    Protected Overloads Sub Initializing(ByRef m As Model2,
                                         ByRef vm As ViewModel,
                                         ByRef pim As ProjectInfoModel)
        Model = m
        ViewModel = vm
        ProjectInfo = pim
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model2,
                                         ByRef vm As ViewModel,
                                         ByRef pim As ProjectInfoModel,
                                         ByRef cmcp As ContextModelCheckProxy)
        Initializing(m, vm, pim)

        ' ビューモデルの設定（画面登録など）を行います
        cmcp()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model2,
                                         ByRef vm As ViewModel,
                                         ByRef pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef cmcp As ContextModelCheckProxy)
        Initializing(m, vm, pim)

        ' 自身(ビューモデル）の初期化設定を行います
        ip()

        ' ビューモデルの設定（画面登録など）を行います
        cmcp()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model2,
                                         ByRef vm As ViewModel,
                                         ByRef pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef cmcp As ContextModelCheckProxy,
                                         ByRef cccep As CheckCommandEnabledProxy)
        Initializing(m, vm, pim, ip, cmcp)

        ' コマンド実行可否の設定
        cccep()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model2,
                                         ByRef vm As ViewModel,
                                         ByRef pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef cmcp As ContextModelCheckProxy,
                                         ByRef ahp As AddHandlerProxy)
        Initializing(m, vm, pim, ip, cmcp)

        ' イベントハンドラの登録
        ahp()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model2,
                                         ByRef vm As ViewModel,
                                         ByRef pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef cmcp As ContextModelCheckProxy,
                                         ByRef cccep As CheckCommandEnabledProxy,
                                         ByRef ahp As AddHandlerProxy)
        Initializing(m, vm, pim, ip, cmcp, cccep)

        ' イベントハンドラの登録
        ahp()
    End Sub
End Class
