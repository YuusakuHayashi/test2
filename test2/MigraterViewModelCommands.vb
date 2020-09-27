'Public Class MigraterViewModelCommands
'    '--- ＡＤＤ ------------------------------------------------------------------------'
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

'    ' コマンド有効／無効判定(ＡＤＤ)
'    Private Sub _CheckAddCommandEnabled()
'        Dim b As Boolean : b = True
'        If Me.Conditions.Count > 0 Then
'            For Each bro In Me.Conditions
'                If String.IsNullOrEmpty(bro.FieldName) Then
'                    b = False
'                End If
'                If String.IsNullOrEmpty(bro.FieldValue) Then
'                    b = False
'                End If
'                If Not b Then
'                    Exit For
'                End If
'            Next
'        End If
'        Me._AddCommandEnableFlag = b
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

'    ' コマンド実行(ＡＤＤ)
'    Private Sub _AddCommandExecute(ByVal parameter As Object)
'        Dim occm As ObservableCollection(Of ConditionModel)
'        occm = Me.Conditions
'        If occm.Count > 0 Then
'            occm.Add(New ConditionModel With {.Logic = "Or"})
'        Else
'            occm.Add(New ConditionModel With {.Logic = vbNullString})
'        End If
'        Me.Conditions = occm
'    End Sub

'    ' コマンド有効／無効化(ＡＤＤ)
'    Private Function _AddCommandCanExecute(ByVal parameter As Object) As Boolean
'        Return Me._AddCommandEnableFlag
'    End Function

'    '-----------------------------------------------------------------------------------'
'End Class
