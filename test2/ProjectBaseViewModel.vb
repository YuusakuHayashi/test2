'https://setonaikai1982.com/mvvm-comm/#ViewModel

Imports System.ComponentModel

Public MustInherit Class ProjectBaseViewModel(Of T)
    Inherits ProjectModel(Of T)
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

    Private _ViewModel As ViewModel
    Public Property ViewModel As ViewModel
        Get
            Return _ViewModel
        End Get
        Set(value As ViewModel)
            _ViewModel = value
        End Set
    End Property

    Private _Commands As Object
    Public Property Commands As Object
        Get
            Return _Commands
        End Get
        Set(value As Object)
            _Commands = value
        End Set
    End Property

    Protected Delegate Sub ViewModelInitializingProxy()
    Protected Delegate Sub CheckCommandEnabledProxy()
    Protected Delegate Sub AddHandlerProxy()

    Protected Overridable Sub MyInitializing()
        'Nothing To Do
    End Sub

    Protected Overridable Sub ContextModelCheck()
        'Nothing To Do
    End Sub


    Protected Overloads Sub Initializing(ByRef m As Model,
                                         ByRef vm As ViewModel)
        Model = m
        If Model Is Nothing Then
            Model = New Model
        End If

        ViewModel = vm
        If ViewModel Is Nothing Then
            ViewModel = New ViewModel
        End If

        ' モデルの設定
        ' 各ビューモデルは、個別の設定をオーバライドします
        MyInitializing()

        ' ビューモデルの設定
        ' 各ビューモデルは、個別の設定をオーバライドして使用します
        ContextModelCheck()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model,
                                         ByRef vm As ViewModel,
                                         ByRef cccep As CheckCommandEnabledProxy)
        Initializing(m, vm)

        ' コマンド実行可否の設定
        cccep()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model,
                                         ByRef vm As ViewModel,
                                         ByRef ahp As AddHandlerProxy)
        Initializing(m, vm)

        ' イベントハンドラの登録
        ahp()
    End Sub

    Protected Overloads Sub Initializing(ByRef m As Model,
                                         ByRef vm As ViewModel,
                                         ByRef cccep As CheckCommandEnabledProxy,
                                         ByRef ahp As AddHandlerProxy)
        Initializing(m, vm, cccep)

        ' イベントハンドラの登録
        ahp()
    End Sub
End Class
