Public Class MenuViewModel : Inherits BaseViewModel

    Private Const _MENUVIEW_FILE As String = "MenuView.txt"
    Private Const _BATCHMENU As String = "BatchMenu"


    'セーブ用にモデルを保持
    'ビューモデルをテキストにデシリアライズしようとすると
    'Newtonsoft.jsonでは、コマンドがデシリアライズ不可
    Private __Model As MenuModel
    Private Property _Model As MenuModel
        Get
            Return Me.__Model
        End Get
        Set(value As MenuModel)
            Me.__Model = value
        End Set
    End Property


    'この画面に／から遷移するフラグ
    'ウィンドウ側で購読する
    Private _MenuFlag As Boolean
    Public Property MenuFlag As Boolean
        Get
            Return Me._MenuFlag
        End Get
        Set(value As Boolean)
            Me._MenuFlag = value
            RaisePropertyChanged("MenuFlag")
        End Set
    End Property


    Private _BatchMenuCommand As ICommand
    Public ReadOnly Property BatchMenuCommand As ICommand
        Get
            If Me._BatchMenuCommand Is Nothing Then
                Me._BatchMenuCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _BatchMenuCommandExecute,
                    .CanExecuteHandler = AddressOf _BatchMenuCommandCanExecute
                }
                Return Me._BatchMenuCommand
            Else
                Return Me._BatchMenuCommand
            End If
        End Get
    End Property



    '初期時に遷移するメニュー画面
    Private __DefaultMenu As String
    Private Property _DefaultMenu As String
        Get
            Return Me.__DefaultMenu
        End Get
        Set(value As String)
            Me.__DefaultMenu = value
        End Set
    End Property


    '初期時にメニュー画面をスルーするフラグ
    Private Sub _CheckMenuEnableFlag()
        Dim b As Boolean : b = False
        Select Case Me._DefaultMenu
            Case MenuViewModel._BATCHMENU
                b = True
            Case Else
                b = False
        End Select
        Me.MenuFlag = b
    End Sub


    'コマンド実行可否のフラグ
    Private __BatchMenuEnableFlag As Boolean
    Public Property _BatchMenuEnableFlag As Boolean
        Get
            Return Me.__BatchMenuEnableFlag
        End Get
        Set(value As Boolean)
            Me.__BatchMenuEnableFlag = value
            RaisePropertyChanged("BatchMenuEnableFlag")
            CType(BatchMenuCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property


    'バッチ処理画面へ遷移する。および設定ファイルへの記録
    'メニューフラグを変更することで、ウィンドウ側はその変更を購読し、切り替えをする
    Private Sub _BatchMenuCommandExecute(ByVal parameter As Object)
        'Dim pm As New ProjectModel
        'Dim proxy As ProjectCheckProxy


        Me._DefaultMenu = MenuViewModel._BATCHMENU
        Me._Model.DefaultMenu = Me._DefaultMenu


        'フラグをＯＮに
        Call Me._CheckMenuEnableFlag()
        Me._Model.MenuFlag = Me.MenuFlag


        '外部に記録
        'proxy = AddressOf pm.FileCheck
        'Select Case pm.ProjectCheck(proxy, _MENUVIEW_FILE)
        '    Case 0
        '        pm.ModelSave(Of MenuModel)(_MENUVIEW_FILE, Me._Model)
        '    Case 1
        '        pm.FileEstablish(_MENUVIEW_FILE)
        '        pm.ModelSave(Of MenuModel)(_MENUVIEW_FILE, Me._Model)
        '    Case 99
        '    Case 999
        'End Select
    End Sub


    '現状、無条件で許可
    Private Function _BatchMenuCommandCanExecute(ByVal parameter As Object) As Boolean
        Return True
    End Function


    Private Sub _SetDefault()
        Me._DefaultMenu = vbNullString
        Me.MenuFlag = False
    End Sub



    Sub New()
        'Dim pm As New ProjectModel
        'Dim proxy As ProjectCheckProxy
        'Dim proxy2 As ModelLoadProxy(Of MenuModel)

        'proxy = AddressOf pm.FileCheck
        'proxy2 = AddressOf pm.ModelLoad(Of MenuModel)

        'Select Case pm.ProjectCheck(proxy, _MENUVIEW_FILE)
        '    Case 0
        '        Me._Model = proxy2(_MENUVIEW_FILE)
        '        If Me._Model IsNot Nothing Then
        '            Me._DefaultMenu = Me._Model.DefaultMenu
        '            If String.IsNullOrEmpty(Me._DefaultMenu) Then
        '                Me._DefaultMenu = vbNullString
        '            End If

        '            Me.MenuFlag = Me._Model.MenuFlag
        '            If Me._Model.MenuFlag = Nothing Then
        '                Me.MenuFlag = False
        '            End If
        '        Else
        '            Me._Model = New MenuModel
        '            Call Me._SetDefault()
        '        End If
        '    Case Else
        '        Me._Model = New MenuModel
        '        Call Me._SetDefault()
        'End Select
    End Sub
End Class
