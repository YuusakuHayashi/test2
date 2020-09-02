Imports System.Collections.ObjectModel

Public Class ConditionModel : Inherits BaseModel

    Private _FieldName As String
    Public Property FieldName As String
        Get
            Return Me._FieldName
        End Get
        Set(value As String)
            Me._FieldName = value
        End Set
    End Property

    Private _FieldValue As Object
    Public Property FieldValue As Object
        Get
            Return Me._FieldValue
        End Get
        Set(value As Object)
            Me._FieldValue = value
        End Set
    End Property

    Private _ConditionChange As String
    Public Property ConditionChange As String
        Get
            Return Me._ConditionChange
        End Get
        Set(value As String)
            Me._ConditionChange = value
            RaisePropertyChanged("ConditionChange")
        End Set
    End Property


    Private _Conditions As ObservableCollection(Of ConditionModel)
    Public Property Conditions As ObservableCollection(Of ConditionModel)
        Get
            Return Me._Conditions
        End Get
        Set(value As ObservableCollection(Of ConditionModel))
            Me._Conditions = value
            RaisePropertyChanged("Conditions")
        End Set
    End Property

    ' コマンドプロパティ(ＯＲ)
    Private _AddCommand As ICommand
    Public ReadOnly Property AddCommand As ICommand
        Get
            If Me._AddCommand Is Nothing Then
                Me._AddCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _AddCommandExecute,
                    .CanExecuteHandler = AddressOf _AddCommandCanExecute
                }
                Return Me._AddCommand
            Else
                Return Me._AddCommand
            End If
        End Get
    End Property

    ' コマンドプロパティ(ＡＮＤ)
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

    ' コマンド有効／無効判定(ＯＲ)
    Private Sub _CheckAddCommandEnableFlag()
        Me._AddCommandEnableFlag = True
    End Sub

    ' コマンド有効／無効判定(ＡＮＤ)
    Private Sub _CheckDeleteCommandEnableFlag()
        Me._DeleteCommandEnableFlag = True
    End Sub

    'コマンド実行可否のフラグ
    Private __AddCommandEnableFlag As Boolean
    Public Property _AddCommandEnableFlag As Boolean
        Get
            Return Me.__AddCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__AddCommandEnableFlag = value
            RaisePropertyChanged("_AddCommandEnableFlag")
            CType(AddCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    'コマンド実行可否のフラグ
    Private __DeleteCommandEnableFlag As Boolean
    Private disposedValue As Boolean

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

    ' コマンド実行(ＯＲ)
    Private Sub _AddCommandExecute(ByVal parameter As Object)
        If Me.Conditions Is Nothing Then
            Me.Conditions = New ObservableCollection(Of ConditionModel)
        End If
        Me.Conditions.Add(New ConditionModel)
    End Sub

    ' コマンド実行(ＡＮＤ)
    Private Sub _DeleteCommandExecute(ByVal parameter As Object)
        Me.Finalize()
    End Sub

    ' コマンド有効／無効化(ＯＲ)
    Private Function _AddCommandCanExecute(ByVal parameter As Object) As Boolean
        'Return Me._AddCommandEnableFlag
        Return True
    End Function

    ' コマンド有効／無効化(ＡＮＤ)
    Private Function _DeleteCommandCanExecute(ByVal parameter As Object) As Boolean
        'Return Me._DeleteCommandEnableFlag
        Return True
    End Function

    Sub Test()
        MsgBox("hoge")
    End Sub

    Sub New()
    End Sub
End Class
