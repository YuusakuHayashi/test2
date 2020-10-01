'Imports System.ComponentModel
'Imports System.Collections.ObjectModel

Public Class DBExplorerViewModel
    '    Inherits ProjectBaseViewModel(Of DBExplorerViewModel)

    '    ' モデル
    '    Private _Server As ServerModel
    '    Public Property Server As ServerModel
    '        Get
    '            Return Me._Server
    '        End Get
    '        Set(value As ServerModel)
    '            Me._Server = value
    '            RaisePropertyChanged("Server")
    '        End Set
    '    End Property

    '    ' モデル
    '    Private _MenuFolder As MenuFolderModel
    '    Public Property MenuFolder As MenuFolderModel
    '        Get
    '            Return Me._MenuFolder
    '        End Get
    '        Set(value As MenuFolderModel)
    '            Me._MenuFolder = value
    '            RaisePropertyChanged("MenuFolder")
    '        End Set
    '    End Property

    ''--- リロード ----------------------------------------------------------------------'
    '    ' コマンドプロパティ(リロード)
    '    Private _DBLoadCommand As ICommand
    '    Public ReadOnly Property DBLoadCommand As ICommand
    '        Get
    '            If Me._DBLoadCommand Is Nothing Then
    '                Me._DBLoadCommand = New DelegateCommand With {
    '                    .ExecuteHandler = AddressOf _DBLoadCommandExecute,
    '                    .CanExecuteHandler = AddressOf _DBLoadCommandCanExecute
    '                }
    '                Return Me._DBLoadCommand
    '            Else
    '                Return Me._DBLoadCommand
    '            End If
    '        End Get
    '    End Property

    '    ' コマンド有効／無効判定(リロード)
    '    Private Sub _CheckDBReLoadCommandEnabled()
    '        Dim b As Boolean
    '        If String.IsNullOrEmpty(Me.Model.ConnectionString) Then
    '            b = False
    '        Else
    '            b = True
    '        End If
    '        Me._DBLoadEnableFlag = b
    '    End Sub

    '    ' コマンド有効／無効フラグ(リロード)
    '    Private __DBLoadEnableFlag As Boolean
    '    Public Property _DBLoadEnableFlag As Boolean
    '        Get
    '            Return Me.__DBLoadEnableFlag
    '        End Get
    '        Set(value As Boolean)
    '            Me.__DBLoadEnableFlag = value
    '            RaisePropertyChanged("DBLoadEnableFlag")
    '            CType(DBLoadCommand, DelegateCommand).RaiseCanExecuteChanged()
    '        End Set
    '    End Property

    '    ' コマンド実行(リロード)
    '    Private Sub _DBLoadCommandExecute(ByVal parameter As Object)
    '        Call Me.Model.ReLoadServer()
    '        Me.Server = Me.Model.Server
    '    End Sub

    '    ' コマンド有効／無効化(リロード)
    '    Private Function _DBLoadCommandCanExecute(ByVal parameter As Object) As Boolean
    '        Return Me._DBLoadEnableFlag
    '    End Function
    ''-----------------------------------------------------------------------------------'


    ''--- セーブ ------------------------------------------------------------------------'
    '    ' コマンドプロパティ(セーブ)
    '    Private _SaveCommand As ICommand
    '    Public ReadOnly Property SaveCommand As ICommand
    '        Get
    '            If Me._SaveCommand Is Nothing Then
    '                Me._SaveCommand = New DelegateCommand With {
    '                    .ExecuteHandler = AddressOf _SaveCommandExecute,
    '                    .CanExecuteHandler = AddressOf _SaveCommandCanExecute
    '                }
    '                Return Me._SaveCommand
    '            Else
    '                Return Me._SaveCommand
    '            End If
    '        End Get
    '    End Property

    '    ' コマンド有効／無効判定(セーブ)
    '    Private Sub _CheckSaveCommandEnabled()
    '        Dim b As Boolean
    '        b = True
    '        Me._SaveEnableFlag = b
    '    End Sub

    '    ' コマンド有効／無効のフラグ(セーブ)
    '    Private __SaveEnableFlag As Boolean
    '    Public Property _SaveEnableFlag As Boolean
    '        Get
    '            Return Me.__SaveEnableFlag
    '        End Get
    '        Set(value As Boolean)
    '            Me.__SaveEnableFlag = value
    '            RaisePropertyChanged("SaveEnableFlag")
    '            CType(SaveCommand, DelegateCommand).RaiseCanExecuteChanged()
    '        End Set
    '    End Property

    '    ' コマンド実行(セーブ)
    '    Private Sub _SaveCommandExecute(ByVal parameter As Object)
    '        Call ModelSave(Of Model)(Model.CurrentModelJson, Me.Model)
    '    End Sub

    '    ' コマンド有効／無効化(セーブ)
    '    Private Function _SaveCommandCanExecute(ByVal parameter As Object) As Boolean
    '        Return Me._SaveEnableFlag
    '    End Function
    ''-----------------------------------------------------------------------------------'


    '    ' このメソッドはInitializing実行時に実行されます
    '    Protected Overrides Sub MyInitializing()
    '        Me.Server = Model.Server

    '        ' イベント購読メソッドの呼び出し方法については現状これで妥協している
    '        ' 他の方法も検討中
    '        Model.MenuFolder.MenuChangeRequestedAddHandler()
    '        Me.MenuFolder = Model.MenuFolder
    '    End Sub


    '    Protected Overrides Sub ContextModelCheck()
    '        ViewModel.SetContext(ViewModel.EXPLORER_VIEW, Me.GetType.Name, Me)
    '    End Sub


    '    ' Initializingメソッド
    '    Sub New(ByRef m As Model,
    '            ByRef vm As ViewModel)
    '        Dim ccep(1) As CheckCommandEnabledProxy
    '        Dim ccep2 As CheckCommandEnabledProxy

    '        ccep(0) = AddressOf Me._CheckDBReLoadCommandEnabled
    '        ccep(1) = AddressOf Me._CheckSaveCommandEnabled
    '        ccep2 = [Delegate].Combine(ccep)

    '        Initializing(m, vm, ccep2)
    '    End Sub
End Class
