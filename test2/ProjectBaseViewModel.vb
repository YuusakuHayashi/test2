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

        ' ビューフラグの設定
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


    Public Sub MemberCheck()
        '
    End Sub
End Class

'Public Class DelegateCommand : Implements ICommand

'    Public Property CanExecuteDelegater As Func(Of Object, Boolean)
'    Public Property ExecuteDelegater As Action(Of String)

'    Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

'    Public Sub RaiseCanExecuteChanged()
'        RaiseEvent CanExecuteChanged(Me, Nothing)
'    End Sub

'    Public Sub Execute(parameter As Object) Implements ICommand.Execute
'        Dim d = ExecuteDelegater
'        If d <> Nothing Then
'            d(parameter)
'        End If
'    End Sub

'    Public Function CanExecute(ByVal parameter As Object) As Boolean Implements ICommand.CanExecute
'        Dim d = CanExecuteDelegater
'        Return IIf(d = Nothing, True, d(parameter))
'    End Function
'End Class


'Public Class PropertyChangedEventListener : Implements IDisposable

'#Region "IDisposable Support"

'    '2020-07-28 : 購読 ---------------------------------------'
'    '参考 : https://qiita.com/nossey/items/7c415799bc6fda45f94e

'    Private _Source As INotifyPropertyChanged
'    Private _Handler As PropertyChangedEventHandler


'    Public Sub New(ByRef inpc As INotifyPropertyChanged,
'                               ByRef pceh As PropertyChangedEventHandler)
'        _Source = inpc
'        _Handler = pceh
'        AddHandler _Source.PropertyChanged, _Handler
'    End Sub
'    '---------------------------------------------------------'


'    Private disposedValue As Boolean ' 重複する呼び出しを検出するには

'    ' IDisposable
'    Protected Overridable Sub Dispose(disposing As Boolean)
'        If Not disposedValue Then
'            If disposing Then
'                ' TODO: マネージ状態を破棄します (マネージ オブジェクト)。

'                '2020-07-28 : 購読解除 -------------------------------------'
'                If _Source IsNot Nothing Then
'                    If _Handler IsNot Nothing Then
'                        RemoveHandler _Source.PropertyChanged, _Handler
'                    End If
'                End If
'                '-----------------------------------------------------------'
'            End If

'            ' TODO: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下の Finalize() をオーバーライドします。
'            ' TODO: 大きなフィールドを null に設定します。
'        End If
'        disposedValue = True
'    End Sub

'    ' TODO: 上の Dispose(disposing As Boolean) にアンマネージ リソースを解放するコードが含まれる場合にのみ Finalize() をオーバーライドします。
'    'Protected Overrides Sub Finalize()
'    '    ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(disposing As Boolean) に記述します。
'    '    Dispose(False)
'    '    MyBase.Finalize()
'    'End Sub

'    ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
'    Public Sub Dispose() Implements IDisposable.Dispose
'        ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(disposing As Boolean) に記述します。
'        Dispose(True)
'        ' TODO: 上の Finalize() がオーバーライドされている場合は、次の行のコメントを解除してください。
'        ' GC.SuppressFinalize(Me)
'    End Sub
'#End Region
'End Class
