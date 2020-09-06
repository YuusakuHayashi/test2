'Imports System.Collections.ObjectModel

'Public Class HeaderConditionModel : Inherits BaseModel

'    Private _Conditions As ObservableCollection(Of ConditionModel)
'    Public Property Conditions As ObservableCollection(Of ConditionModel)
'        Get
'            Return Me._Conditions
'        End Get
'        Set(value As ObservableCollection(Of ConditionModel))
'            Me._Conditions = value
'            RaisePropertyChanged("Conditions")
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

'    Private _AddRequest As Boolean
'    Public Property AddRequest As Boolean
'        Get
'            Return Me._AddRequest
'        End Get
'        Set(value As Boolean)
'            Me._AddRequest = value
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

'    ' コマンド有効／無効判定(ＡＤＤ)
'    Private Sub _CheckAddCommandEnableFlag()

'        Dim b As Boolean
'        b = True
'        'Dim func As Func(Of ConditionModel, Boolean)
'        'func = Function(ByVal cm As ConditionModel)
'        '           Dim lb As Boolean
'        '           lb = True
'        '           If String.IsNullOrEmpty(cm.FieldName) Then
'        '               lb = False
'        '           End If
'        '           If String.IsNullOrEmpty(cm.FieldValue) Then
'        '               lb = False
'        '           End If
'        '           Return lb
'        '       End Function

'        'b = func(Me)

'        'If b Then
'        '    If Me.Conditions IsNot Nothing Then
'        '        For Each child In Me.Conditions
'        '            b = func(child)
'        '            If Not b Then
'        '                Exit For
'        '            End If
'        '        Next
'        '    End If
'        'End If

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
'        Me.AddRequest = True
'        ConditionChangedEventListener.Instance.RaiseHeaderAddRequested(Me)
'        'If Me.Conditions Is Nothing Then
'        '    Me.Conditions = New ObservableCollection(Of ConditionModel)
'        '    Me.Conditions.Add(New ConditionModel With {.Logic = "And"})
'        'Else
'        '    Me.Conditions.Add(New ConditionModel With {.Logic = "Or"})
'        'End If
'        'Call Me._CheckAddCommandEnableFlag()
'    End Sub

'    ' コマンド実行(ＤＥＬＥＴＥ)
'    Private Sub _DeleteCommandExecute(ByVal parameter As Object)
'        Me.DeleteRequest = True
'        ConditionChangedEventListener.Instance.RaiseHeaderDeleteRequested(Me)
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

'    Sub DeleteRequestReview(ByVal sender As Object, ByVal e As EventArgs)
'        Call Me.DeleteAccept()
'    End Sub

'    Sub DeleteAccept()
'        Dim occm As ObservableCollection(Of ConditionModel)
'        occm = DeletedCondition(Me.Conditions)
'        Me.Conditions = occm
'    End Sub

'    Public Function DeletedCondition(
'        ByVal occm As ObservableCollection(Of ConditionModel)) As ObservableCollection(Of ConditionModel)
'        Dim lcm As New List(Of ConditionModel)

'        For Each child In occm
'            If child.DeleteRequest Then
'                lcm.Add(child)
'            Else
'                If child.Conditions IsNot Nothing Then
'                    child.Conditions = DeletedCondition(child.Conditions)
'                End If
'            End If
'        Next

'        For Each child In lcm
'            occm.Remove(child)
'        Next

'        DeletedCondition = occm
'    End Function

'    Sub New()
'        ' コマンドフラグの有効／無効化
'        Call Me._CheckAddCommandEnableFlag()

'        AddHandler ConditionChangedEventListener.Instance.DeleteRequested, AddressOf Me.DeleteRequestReview
'    End Sub

'    Protected Overrides Sub Finalize()
'        RemoveHandler ConditionChangedEventListener.Instance.DeleteRequested, AddressOf Me.DeleteRequestReview
'        MyBase.Finalize()
'    End Sub
'End Class
