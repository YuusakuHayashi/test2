Public Class TabItemModel : Inherits BaseModel(Of Object)
    Private Property _Name As String
    Public Property Name As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
        End Set
    End Property

    Private Property _Content As Object
    Public Property Content As Object
        Get
            Return Me._Content
       End Get
        Set(value As Object)
            Me._Content = value
        End Set
    End Property

    ' コマンドプロパティ（接続確認）
    Private _TabCloseCommand As ICommand
    Public ReadOnly Property TabCloseCommand As ICommand
        Get
            If Me._TabCloseCommand Is Nothing Then
                Me._TabCloseCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _TabCloseCommandExecute,
                    .CanExecuteHandler = AddressOf _TabCloseCommandCanExecute
                }
                Return Me._TabCloseCommand
            Else
                Return Me._TabCloseCommand
            End If
        End Get
    End Property

    'コマンド実行可否のチェック（接続確認）
    Private Sub _CheckTabCloseCommandEnabled()
        Dim b = True
        Me._TabCloseCommandEnableFlag = b
    End Sub

    ' コマンド実行可否のフラグ（接続確認）
    Private __TabCloseCommandEnableFlag As Boolean
    Public Property _TabCloseCommandEnableFlag As Boolean
        Get
            Return Me.__TabCloseCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__TabCloseCommandEnableFlag = value
            RaisePropertyChanged("_TabCloseCommandEnableFlag")
            CType(TabCloseCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    ' コマンド実行（接続確認）
    Private Sub _TabCloseCommandExecute(ByVal parameter As Object)
        Call DelegateEventListener.Instance.RaiseTabCloseRequested(Me)
    End Sub

    Private Function _TabCloseCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._TabCloseCommandEnableFlag
    End Function

    Public Sub Initialize()
        Call _CheckTabCloseCommandEnabled()
    End Sub
End Class
