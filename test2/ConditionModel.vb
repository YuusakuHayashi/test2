'Imports System.Collections.ObjectModel

Public Class ConditionModel : Inherits BaseModel(Of ConnectionModel)
    '    Private _Logic As String
    '    Public Property Logic As String
    '        Get
    '            Return Me._Logic
    '        End Get
    '        Set(value As String)
    '            Me._Logic = value
    '            RaisePropertyChanged("Logic")
    '        End Set
    '    End Property

    '    Private _FieldName As String
    '    Public Property FieldName As String
    '        Get
    '            Return Me._FieldName
    '        End Get
    '        Set(value As String)
    '            Me._FieldName = value
    '            Call Me._CheckAddCommandEnableFlag()
    '            DelegateEventListener.Instance.RaiseChildrenUpdated(Me)
    '        End Set
    '    End Property

    '    Private _FieldValue As Object
    '    Public Property FieldValue As Object
    '        Get
    '            Return Me._FieldValue
    '        End Get
    '        Set(value As Object)
    '            Me._FieldValue = value
    '            Call Me._CheckAddCommandEnableFlag()
    '            DelegateEventListener.Instance.RaiseChildrenUpdated(Me)
    '        End Set
    '    End Property

    '    Private _DeleteRequest As Boolean
    '    Public Property DeleteRequest As Boolean
    '        Get
    '            Return Me._DeleteRequest
    '        End Get
    '        Set(value As Boolean)
    '            Me._DeleteRequest = value
    '        End Set
    '    End Property


    '    Private _Conditions As ObservableCollection(Of ConditionModel)
    '    Public Property Conditions As ObservableCollection(Of ConditionModel)
    '        Get
    '            Return Me._Conditions
    '        End Get
    '        Set(value As ObservableCollection(Of ConditionModel))
    '            Me._Conditions = value
    '            RaisePropertyChanged("Conditions")
    '            Call Me._CheckAddCommandEnableFlag()
    '        End Set
    '    End Property

    '    ' コマンドプロパティ(ＡＤＤ)
    '    Private _AddCommand As ICommand
    '    Public ReadOnly Property AddCommand As ICommand
    '        Get
    '            If Me._AddCommand Is Nothing Then
    '                Me._AddCommand = New DelegateCommand With {
    '                    .ExecuteHandler = AddressOf _AddCommandExecute,
    '                    .CanExecuteHandler = AddressOf _AddCommandCanExecute
    '                }
    '                Return Me._AddCommand
    '            Else
    '                Return Me._AddCommand
    '            End If
    '        End Get
    '    End Property

    '    ' コマンドプロパティ(ＤＥＬＥＴＥ)
    '    Private _DeleteCommand As ICommand
    '    Public ReadOnly Property DeleteCommand As ICommand
    '        Get
    '            If Me._DeleteCommand Is Nothing Then
    '                Me._DeleteCommand = New DelegateCommand With {
    '                    .ExecuteHandler = AddressOf _DeleteCommandExecute,
    '                    .CanExecuteHandler = AddressOf _DeleteCommandCanExecute
    '                }
    '                Return Me._DeleteCommand
    '            Else
    '                Return Me._DeleteCommand
    '            End If
    '        End Get
    '    End Property

    '    ' コマンド有効／無効判定(ＡＤＤ)
    '    Private Sub _CheckAddCommandEnableFlag()

    '        Dim b As Boolean
    '        Dim func As Func(Of ConditionModel, Boolean)
    '        func = Function(ByVal cm As ConditionModel)
    '                   Dim lb As Boolean
    '                   lb = True
    '                   If String.IsNullOrEmpty(cm.FieldName) Then
    '                       lb = False
    '                   End If
    '                   If String.IsNullOrEmpty(cm.FieldValue) Then
    '                       lb = False
    '                   End If
    '                   Return lb
    '               End Function

    '        b = func(Me)

    '        If b Then
    '            If Me.Conditions IsNot Nothing Then
    '                If Me.Conditions.Count > 0 Then
    '                    For Each child In Me.Conditions
    '                        b = func(child)
    '                        If Not b Then
    '                            Exit For
    '                        End If
    '                    Next
    '                End If
    '            End If
    '        End If

    '        Me._AddCommandEnableFlag = b
    '    End Sub

    '    ' コマンド有効／無効判定(ＤＥＬＥＴＥ)
    '    Private Sub _CheckDeleteCommandEnableFlag()
    '        Me._DeleteCommandEnableFlag = True
    '    End Sub

    '    'コマンド実行可否のフラグ(ＡＤＤ)
    '    Private __AddCommandEnableFlag As Boolean
    '    Public Property _AddCommandEnableFlag As Boolean
    '        Get
    '            Return Me.__AddCommandEnableFlag
    '        End Get
    '        Set(value As Boolean)
    '            Me.__AddCommandEnableFlag = value
    '            RaisePropertyChanged("_AddCommandEnableFlag")
    '            CType(AddCommand, DelegateCommand).RaiseCanExecuteChanged()
    '        End Set
    '    End Property

    '    'コマンド実行可否のフラグ(ＤＥＬＥＴＥ)
    '    Private __DeleteCommandEnableFlag As Boolean

    '    Public Property _DeleteCommandEnableFlag As Boolean
    '        Get
    '            Return Me.__DeleteCommandEnableFlag
    '        End Get
    '        Set(value As Boolean)
    '            Me.__DeleteCommandEnableFlag = value
    '            RaisePropertyChanged("_DeleteCommandEnableFlag")
    '            CType(DeleteCommand, DelegateCommand).RaiseCanExecuteChanged()
    '        End Set
    '    End Property

    '    ' コマンド実行(ＡＤＤ)
    '    Private Sub _AddCommandExecute(ByVal parameter As Object)
    '        Dim occm As ObservableCollection(Of ConditionModel)
    '        occm = Me.Conditions
    '        If occm Is Nothing Then
    '            occm = New ObservableCollection(Of ConditionModel)
    '            occm.Add(New ConditionModel With {.Logic = "And"})
    '        Else
    '            occm.Add(New ConditionModel With {.Logic = "Or"})
    '        End If
    '        Me.Conditions = occm
    '    End Sub

    '    ' コマンド実行(ＤＥＬＥＴＥ)
    '    Private Sub _DeleteCommandExecute(ByVal parameter As Object)
    '        Me.DeleteRequest = True
    '        DelegateEventListener.Instance.RaiseDeleteRequested(Me)
    '    End Sub

    '    ' コマンド有効／無効化(ＡＤＤ)
    '    Private Function _AddCommandCanExecute(ByVal parameter As Object) As Boolean
    '        Return Me._AddCommandEnableFlag
    '    End Function

    '    ' コマンド有効／無効化(ＤＥＬＥＴＥ)
    '    Private Function _DeleteCommandCanExecute(ByVal parameter As Object) As Boolean
    '        'Return Me._DeleteCommandEnableFlag
    '        Return True
    '    End Function

    '    ' 更新通知処理
    '    '' 全ての要素が変更通知をトラップしてしまうので、ここで変更通知を出した子の親のみに限定する
    '    Private Sub _ChildrenUpdatedReview(ByVal sender As Object, ByVal e As EventArgs)
    '        Dim cm As ConditionModel
    '        cm = CType(sender, ConditionModel)

    '        If Me.Conditions IsNot Nothing Then
    '            If Me.Conditions.Contains(cm) Then
    '                Call Me._CheckAddCommandEnableFlag()
    '            End If
    '        End If
    '    End Sub

    '    Sub New()
    '        RemoveHandler _
    '            DelegateEventListener.Instance.ChildrenUpdated,
    '            AddressOf Me._ChildrenUpdatedReview
    '        AddHandler _
    '            DelegateEventListener.Instance.ChildrenUpdated,
    '            AddressOf Me._ChildrenUpdatedReview
    '    End Sub
End Class
