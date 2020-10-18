Public Class ViewItemModel
    Inherits BaseViewModel

    Private _BoxVisibility As Visibility
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

    Private _ViewType As String
    Public Property ViewType As String
        Get
            Return Me._ViewType
        End Get
        Set(value As String)
            Me._ViewType = value
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
    Public Property IsSelected As Boolean
        Get
            Return Me._IsSelected
        End Get
        Set(value As Boolean)
            Me._IsSelected = value
            If value Then
                Call DelegateEventListener.Instance.RaiseOpenViewRequested(Me)
            End If
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

    ' コマンドプロパティ（接続確認）
    Private _ChangeAliasCommand As ICommand
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

    Public Sub Initialize()
        Call _CheckChangeAliasCommandEnabled()
    End Sub
End Class
