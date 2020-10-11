﻿Imports System.ComponentModel
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

    Private _MyModel As T
    Public Property MyModel As T
        Get
            Return _MyModel
        End Get
        Set(value As T)
            _MyModel = value
        End Set
    End Property


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


    ' ビューモデル初期化時のメインメソッド
    ' 共有のモデル・ビューモデル・アプリケーション情報、プロジェクト情報をセットする
    ' さらに各ビューモデルに対応するメインモデルもセットする
    Protected Overloads Sub Initializing(ByVal m As Model,
                                         ByVal vm As ViewModel,
                                         ByVal adm As AppDirectoryModel,
                                         ByVal pim As ProjectInfoModel)
        Dim obj As Object

        Model = m
        ViewModel = vm
        AppInfo = adm
        ProjectInfo = pim

        ' ビューモデルに対応するメインモデルをセットする
        If Model.DataDictionary IsNot Nothing Then
            obj = Model.DataDictionary((New T).GetType.Name)
            Select Case obj.GetType
                Case (New Object).GetType
                    Me.MyModel = CType(obj, T)
                Case (New JObject).GetType
                    ' Ｊｓｏｎからロードした場合は、JObject型になっている
                    Me.MyModel = obj.ToObject(Of T)
                Case (New T).GetType
                    Me.MyModel = obj
            End Select
        End If
    End Sub

    Protected Overloads Sub Initializing(ByVal m As Model,
                                         ByVal vm As ViewModel,
                                         ByVal adm As AppDirectoryModel,
                                         ByVal pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy)
        Call Initializing(m, vm, adm, pim)

        ' 自身(ビューモデル）の初期化設定を行います
        Call ip()
    End Sub

    Protected Overloads Sub Initializing(ByVal m As Model,
                                         ByVal vm As ViewModel,
                                         ByVal adm As AppDirectoryModel,
                                         ByVal pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef ccep As CheckCommandEnabledProxy)
        Call Initializing(m, vm, adm, pim, ip)

        ' コマンド実行可否の設定
        Call ccep()
    End Sub

    Protected Overloads Sub Initializing(ByVal m As Model,
                                         ByVal vm As ViewModel,
                                         ByVal adm As AppDirectoryModel,
                                         ByVal pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef ahp As AddHandlerProxy)
        Call Initializing(m, vm, adm, pim, ip)

        ' イベントハンドラの登録
        Call ahp()
    End Sub

    Protected Overloads Sub Initializing(ByVal m As Model,
                                         ByVal vm As ViewModel,
                                         ByVal adm As AppDirectoryModel,
                                         ByVal pim As ProjectInfoModel,
                                         ByRef ip As InitializingProxy,
                                         ByRef ccep As CheckCommandEnabledProxy,
                                         ByRef ahp As AddHandlerProxy)
        Call Initializing(m, vm, adm, pim, ip, ccep)

        ' イベントハンドラの登録
        Call ahp()
    End Sub
End Class