Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ViewItemModel
    Inherits BaseViewModel

    Private _IconFileName As String
    Public Property IconFileName As String
        Get
            Return _IconFileName
        End Get
        Set(value As String)
            _IconFileName = value
        End Set
    End Property

    Private _Icon As BitmapImage
    <JsonIgnore>
    Public Property [Icon] As BitmapImage
        Get
            Return _Icon
        End Get
        Set(value As BitmapImage)
            _Icon = value
        End Set
    End Property

    Private _BoxVisibility As Visibility
    <JsonIgnore>
    Public Property BoxVisibility As Visibility
        Get
            Return Me._BoxVisibility
        End Get
        Set(value As Visibility)
            Me._BoxVisibility = value
            If value = Visibility.Visible Then
                Me.BlockVisibility = Visibility.Collapsed
            End If
            RaisePropertyChanged("BlockVisibility")
        End Set
    End Property

    Private _BlockVisibility As Visibility
    <JsonIgnore>
    Public Property BlockVisibility As Visibility
        Get
            Return Me._BlockVisibility
        End Get
        Set(value As Visibility)
            Me._BlockVisibility = value
            If value = Visibility.Visible Then
                Me.BoxVisibility = Visibility.Collapsed
            End If
            RaisePropertyChanged("BoxVisibility")
        End Set
    End Property

    Private _Name As String
    Public Property Name As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
        End Set
    End Property

    Private _Alias As String
    Public Property [Alias] As String
        Get
            If String.IsNullOrEmpty(Me._Alias) Then
                Me._Alias = Me.Name
            End If
            Return Me._Alias
        End Get
        Set(value As String)
            Me._Alias = value
            If BoxVisibility = Visibility.Visible Then
                BlockVisibility = Visibility.Visible
            End If
            RaisePropertyChanged("Alias")
        End Set
    End Property

    Private _ModelName As String
    Public Property ModelName As String
        Get
            Return Me._ModelName
        End Get
        Set(value As String)
            Me._ModelName = value
        End Set
    End Property

    Private _ViewType As String
    Public Property ViewType As String
        Get
            Return Me._ViewType
        End Get
        Set(value As String)
            Me._ViewType = value
        End Set
    End Property

    Private _FrameType As String
    Public Property FrameType As String
        Get
            Return Me._FrameType
        End Get
        Set(value As String)
            Me._FrameType = value
        End Set
    End Property

    Private _LayoutType As String
    Public Property LayoutType As String
        Get
            Return Me._LayoutType
        End Get
        Set(value As String)
            Me._LayoutType = value
        End Set
    End Property

    Private _OpenState As Boolean
    Public Property OpenState As Boolean
        Get
            Return Me._OpenState
        End Get
        Set(value As Boolean)
            Me._OpenState = value
        End Set
    End Property

    Private _Content As Object
    <JsonIgnore>
    Public Property Content As Object
        Get
            Return Me._Content
        End Get
        Set(value As Object)
            Me._Content = value
        End Set
    End Property

    'Private _IsExpand As Boolean
    'Public Property IsExpand As Boolean
    '    Get
    '        Return Me._IsExpand
    '    End Get
    '    Set(value As Boolean)
    '        Me._IsExpand = value
    '    End Set
    'End Property

    Private _IsSelected As Boolean
    <JsonIgnore>
    Public Property IsSelected As Boolean
        Get
            Return Me._IsSelected
        End Get
        Set(value As Boolean)
            Me._IsSelected = value
            'If value Then
            '    Call DelegateEventListener.Instance.RaiseOpenViewRequested(Me)
            'End If
        End Set
    End Property

    'Private _IsReadOnly As Boolean
    'Public Property IsReadOnly As Boolean
    '    Get
    '        Return Me._IsReadOnly
    '    End Get
    '    Set(value As Boolean)
    '        Me._IsReadOnly = value
    '        RaisePropertyChanged("IsReadOnly")
    '    End Set
    'End Property

    '---------------------------------------------------------------------------------------------'
    ' コマンドプロパティ（接続確認）
    Private _ChangeAliasCommand As ICommand
    <JsonIgnore>
    Public ReadOnly Property ChangeAliasCommand As ICommand
        Get
            If Me._ChangeAliasCommand Is Nothing Then
                Me._ChangeAliasCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ChangeAliasCommandExecute,
                    .CanExecuteHandler = AddressOf _ChangeAliasCommandCanExecute
                }
                Return Me._ChangeAliasCommand
            Else
                Return Me._ChangeAliasCommand
            End If
        End Get
    End Property

    'コマンド実行可否のチェック（接続確認）
    Private Sub _CheckChangeAliasCommandEnabled()
        Dim b As Boolean : b = True
        Me._ChangeAliasCommandEnableFlag = b
    End Sub

    ' コマンド実行可否のフラグ（接続確認）
    Private __ChangeAliasCommandEnableFlag As Boolean
    <JsonIgnore>
    Public Property _ChangeAliasCommandEnableFlag As Boolean
        Get
            Return Me.__ChangeAliasCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ChangeAliasCommandEnableFlag = value
            RaisePropertyChanged("_ChangeAliasCommandEnableFlag")
            CType(ChangeAliasCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    ' コマンド実行（接続確認）
    Private Sub _ChangeAliasCommandExecute(ByVal parameter As Object)
        Me.BoxVisibility = Visibility.Visible
    End Sub

    Private Function _ChangeAliasCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ChangeAliasCommandEnableFlag
    End Function
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    ' 廃止
    ' コマンドプロパティ（ビュー削除）
    'Private _DeleteViewCommand As ICommand
    '<JsonIgnore>
    'Public ReadOnly Property DeleteViewCommand As ICommand
    '    Get
    '        If Me._DeleteViewCommand Is Nothing Then
    '            Me._DeleteViewCommand = New DelegateCommand With {
    '                .ExecuteHandler = AddressOf _DeleteViewCommandExecute,
    '                .CanExecuteHandler = AddressOf _DeleteViewCommandCanExecute
    '            }
    '            Return Me._DeleteViewCommand
    '        Else
    '            Return Me._DeleteViewCommand
    '        End If
    '    End Get
    'End Property

    ''コマンド実行可否のチェック（ビュー削除）
    'Private Sub _CheckDeleteViewCommandEnabled()
    '    Dim b As Boolean : b = True
    '    Me._DeleteViewCommandEnableFlag = b
    'End Sub

    '' コマンド実行可否のフラグ（ビュー削除）
    'Private __DeleteViewCommandEnableFlag As Boolean
    '<JsonIgnore>
    'Public Property _DeleteViewCommandEnableFlag As Boolean
    '    Get
    '        Return Me.__DeleteViewCommandEnableFlag
    '    End Get
    '    Set(value As Boolean)
    '        Me.__DeleteViewCommandEnableFlag = value
    '        RaisePropertyChanged("_DeleteViewCommandEnableFlag")
    '        CType(DeleteViewCommand, DelegateCommand).RaiseCanExecuteChanged()
    '    End Set
    'End Property

    '' コマンド実行（ビュー削除）
    'Private Sub _DeleteViewCommandExecute(ByVal parameter As Object)
    '    Call DelegateEventListener.Instance.RaiseDeleteViewRequested(Me)
    'End Sub

    'Private Function _DeleteViewCommandCanExecute(ByVal parameter As Object) As Boolean
    '    Return Me._DeleteViewCommandEnableFlag
    'End Function
    '---------------------------------------------------------------------------------------------'

    Public Sub Initialize()
        Call _CheckChangeAliasCommandEnabled()
        Me.BlockVisibility = Visibility.Visible
    End Sub

    Public Sub New()
        Call Initialize()
    End Sub
End Class
