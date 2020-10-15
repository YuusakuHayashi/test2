Public Class MigrateModel : Inherits BaseModel(Of MigrateModel)

    Private _Logic As String
    Public Property Logic As String
        Get
            Return Me._Logic
        End Get
        Set(value As String)
            Me._Logic = value
            RaisePropertyChanged("Logic")
        End Set
    End Property

    Private _FieldName As String
    Public Property FieldName As String
        Get
            Return Me._FieldName
        End Get
        Set(value As String)
            Me._FieldName = value
            DelegateEventListener.Instance.RaiseMigrateConditionUpdated(Me)
        End Set
    End Property

    Private _FieldValue As Object
    Public Property FieldValue As Object
        Get
            Return Me._FieldValue
        End Get
        Set(value As Object)
            Me._FieldValue = value
            DelegateEventListener.Instance.RaiseMigrateConditionUpdated(Me)
        End Set
    End Property

    Private _DeleteRequest As Boolean
    Public Property DeleteRequest As Boolean
        Get
            Return Me._DeleteRequest
        End Get
        Set(value As Boolean)
            Me._DeleteRequest = value
        End Set
    End Property

    ' コマンドプロパティ(ＤＥＬＥＴＥ)
    Private _DeleteCommand As ICommand
    Public ReadOnly Property DeleteCommand As ICommand
        Get
            If Me._DeleteCommand Is Nothing Then
                Me._DeleteCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _DeleteCommandExecute,
                    .CanExecuteHandler = AddressOf _DeleteCommandCanExecute
                }
                Return Me._DeleteCommand
            Else
                Return Me._DeleteCommand
            End If
        End Get
    End Property

    ' コマンド有効／無効判定(ＤＥＬＥＴＥ)
    Private Sub _CheckDeleteCommandEnableFlag()
        Dim b As Boolean
        b = True
        'If Me.Logic = vbNullString Then
        '    b = False
        'End If
        Me._DeleteCommandEnableFlag = b
    End Sub

    'コマンド実行可否のフラグ(ＤＥＬＥＴＥ)
    Private __DeleteCommandEnableFlag As Boolean
    Public Property _DeleteCommandEnableFlag As Boolean
        Get
            Return Me.__DeleteCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__DeleteCommandEnableFlag = value
            RaisePropertyChanged("_DeleteCommandEnableFlag")
            CType(DeleteCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    ' コマンド実行(ＤＥＬＥＴＥ)
    Private Sub _DeleteCommandExecute(ByVal parameter As Object)
        Me.DeleteRequest = True
        DelegateEventListener.Instance.RaiseMigrateConditionDeleteRequested(Me)
    End Sub

    ' コマンド有効／無効化(ＤＥＬＥＴＥ)
    Private Function _DeleteCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._DeleteCommandEnableFlag
    End Function

    Sub New()
        Call Me._CheckDeleteCommandEnableFlag()
    End Sub
End Class
