﻿Imports System.Collections.ObjectModel

Public Class MigraterViewModel : Inherits ViewModel
    ' モデル
    Private __Model As Model
    Private Property _Model As Model
        Get
            Return Me.__Model
        End Get
        Set(value As Model)
            Me.__Model = value
            RaisePropertyChanged("Model")
        End Set
    End Property

    ' 条件
    Private _Conditions As ObservableCollection(Of ConditionModel)
    Public Property Conditions As ObservableCollection(Of ConditionModel)
        Get
            Return Me._Conditions
        End Get
        Set(value As ObservableCollection(Of ConditionModel))
            Me._Conditions = value
            RaisePropertyChanged("Conditions")
            Call Me._CheckAddCommandEnableFlag()
        End Set
    End Property

    ' 移行後
    Private _Migrates As ObservableCollection(Of MigrateModel)
    Public Property Migrates As ObservableCollection(Of MigrateModel)
        Get
            Return Me._Migrates
        End Get
        Set(value As ObservableCollection(Of MigrateModel))
            Me._Migrates = value
            RaisePropertyChanged("Migrates")
            Call Me._CheckMigrateConditionAddCommandEnableFlag()
        End Set
    End Property



    ' 削除要求のチェック
    Sub DeleteRequestReview(ByVal sender As Object, ByVal e As EventArgs)
        Call Me.DeleteAccept()
    End Sub


    ' 削除要求受理
    Sub DeleteAccept()
        Dim occm As ObservableCollection(Of ConditionModel)
        occm = DeletedCondition(Me.Conditions)

        If occm.Count > 0 Then
            occm(0).Logic = vbNullString
        End If
        Me.Conditions = occm
    End Sub


    Public Function DeletedCondition(
        ByVal occm As ObservableCollection(Of ConditionModel)) As ObservableCollection(Of ConditionModel)
        Dim lcm As New List(Of ConditionModel)

        For Each child In occm
            If child.DeleteRequest Then
                lcm.Add(child)
            Else
                If child.Conditions IsNot Nothing Then
                    child.Conditions = DeletedCondition(child.Conditions)
                End If
            End If
        Next

        For Each child In lcm
            occm.Remove(child)
        Next

        DeletedCondition = occm
    End Function

    ' 削除要求のチェック(移行後)
    Sub _MigrateConditionDeleteRequestReview(ByVal sender As Object, ByVal e As EventArgs)
        Call Me._MigrateConditionDeleteAccept()
    End Sub

    ' 削除要求受理(移行後)
    Sub _MigrateConditionDeleteAccept()
        Dim ocmm As ObservableCollection(Of MigrateModel)
        ocmm = Me.Migrates
        For Each bro In ocmm
            If bro.DeleteRequest Then
                ocmm.Remove(bro)
                Exit For
            End If
        Next

        If ocmm.Count > 0 Then
            ocmm(0).Logic = vbNullString
        End If
        Me.Migrates = ocmm
    End Sub


    ' コマンドプロパティ(ＡＤＤ)
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

    ' コマンドプロパティ(移行後ＡＤＤ)
    Private _MigrateConditionAddCommand As ICommand
    Public ReadOnly Property MigrateConditionAddCommand As ICommand
        Get
            If Me._MigrateConditionAddCommand Is Nothing Then
                Me._MigrateConditionAddCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _MigrateConditionAddCommandExecute,
                    .CanExecuteHandler = AddressOf _MigrateConditionAddCommandCanExecute
                }
                Return Me._MigrateConditionAddCommand
            Else
                Return Me._MigrateConditionAddCommand
            End If
        End Get
    End Property


    ' コマンドプロパティ(クエリ発行)
    Private _PublishQueryCommand As ICommand
    Public ReadOnly Property PublishQueryCommand As ICommand
        Get
            If Me._PublishQueryCommand Is Nothing Then
                Me._PublishQueryCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _PublishQueryCommandExecute,
                    .CanExecuteHandler = AddressOf _PublishQueryCommandCanExecute
                }
                Return Me._PublishQueryCommand
            Else
                Return Me._PublishQueryCommand
            End If
        End Get
    End Property

    ' コマンド有効／無効判定(ＡＤＤ)
    Private Sub _CheckAddCommandEnableFlag()
        Dim b As Boolean : b = True
        If Me.Conditions.Count > 0 Then
            For Each bro In Me.Conditions
                If String.IsNullOrEmpty(bro.FieldName) Then
                    b = False
                End If
                If String.IsNullOrEmpty(bro.FieldValue) Then
                    b = False
                End If
                If Not b Then
                    Exit For
                End If
            Next
        End If
        Me._AddCommandEnableFlag = b
    End Sub


    ' コマンド有効／無効判定(移行後ＡＤＤ)
    Private Sub _CheckMigrateConditionAddCommandEnableFlag()
        Dim b As Boolean : b = True
        If Me.Migrates.Count > 0 Then
            For Each bro In Me.Migrates
                If String.IsNullOrEmpty(bro.FieldName) Then
                    b = False
                End If
                If String.IsNullOrEmpty(bro.FieldValue) Then
                    b = False
                End If
                If Not b Then
                    Exit For
                End If
            Next
        End If
        Me._MigrateConditionAddCommandEnableFlag = b
    End Sub

    ' コマンド有効／無効判定(クエリ発行)
    Private Sub _CheckPublishQueryCommandEnableFlag()
        Dim b As Boolean : b = True
        Me._PublishQueryCommandEnableFlag = b
    End Sub

    'コマンド実行可否のフラグ(ＡＤＤ)
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

    'コマンド実行可否のフラグ(移行後ＡＤＤ)
    Private __MigrateConditionAddCommandEnableFlag As Boolean
    Public Property _MigrateConditionAddCommandEnableFlag As Boolean
        Get
            Return Me.__MigrateConditionAddCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__MigrateConditionAddCommandEnableFlag = value
            RaisePropertyChanged("_MigrateConditionAddCommandEnableFlag")
            CType(MigrateConditionAddCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    'コマンド実行可否のフラグ(クエリ発行)
    Private __PublishQueryCommandEnableFlag As Boolean
    Public Property _PublishQueryCommandEnableFlag As Boolean
        Get
            Return Me.__PublishQueryCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__PublishQueryCommandEnableFlag = value
            RaisePropertyChanged("_PublishQueryCommandEnableFlag")
            CType(PublishQueryCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    ' コマンド実行(ＡＤＤ)
    Private Sub _AddCommandExecute(ByVal parameter As Object)
        Dim occm As ObservableCollection(Of ConditionModel)
        occm = Me.Conditions
        If occm.Count > 0 Then
            occm.Add(New ConditionModel With {.Logic = "Or"})
        Else
            occm.Add(New ConditionModel With {.Logic = vbNullString})
        End If
        Me.Conditions = occm
    End Sub

    ' コマンド実行(移行後ＡＤＤ)
    Private Sub _MigrateConditionAddCommandExecute(ByVal parameter As Object)
        Dim ocmm As ObservableCollection(Of MigrateModel)
        ocmm = Me.Migrates
        If ocmm.Count > 0 Then
            ocmm.Add(New MigrateModel With {.Logic = "Or"})
        Else
            ocmm.Add(New MigrateModel With {.Logic = vbNullString})
        End If
        Me.Migrates = ocmm
    End Sub

    ' コマンド実行(クエリ発行)
    Private Sub _PublishQueryCommandExecute(ByVal parameter As Object)
        Call Me._AssembleQuery()
    End Sub

    ' コマンド有効／無効化(ＡＤＤ)
    Private Function _AddCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._AddCommandEnableFlag
    End Function

    ' コマンド有効／無効化(移行後ＡＤＤ)
    Private Function _MigrateConditionAddCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._MigrateConditionAddCommandEnableFlag
    End Function

    ' コマンド有効／無効化(クエリ発行)
    Private Function _PublishQueryCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._PublishQueryCommandEnableFlag
    End Function


    Private Sub _AssembleQuery()
        Dim WherePhrase As String
        WherePhrase = "WHERE "
        WherePhrase &= _AssembleQueryBody(Me.Conditions)
        WherePhrase &= ";"
    End Sub

    Private Function _AssembleQueryBody(ByVal occm As ObservableCollection(Of ConditionModel))
        Dim phrase As String : phrase = vbNullString
        For Each child In occm
            If child.Logic = "And" Then
                phrase &= " "
                phrase &= child.Logic
            Else
                phrase &= child.Logic
                phrase &= " ("
            End If
            phrase &= " "
            phrase &= child.FieldName
            phrase &= "="
            phrase &= child.FieldValue
            If child.Conditions IsNot Nothing Then
                If child.Conditions.Count > 0 Then
                    phrase &= Me._AssembleQueryBody(child.Conditions)
                End If
            End If
            If child.Logic = "And" Then
                phrase &= " "
            Else
                phrase &= ") "
            End If
        Next
        _AssembleQueryBody = phrase
    End Function


    Sub New(ByRef m As Model)
        ' モデルの設定
        If m IsNot Nothing Then
        Else
            ' モデルなし
            m = New Model
        End If
        Me._Model = m


        Me.Conditions = New ObservableCollection(Of ConditionModel) From {
            New ConditionModel
        }

        Me.Migrates = New ObservableCollection(Of MigrateModel) From {
            New MigrateModel
        }

        ' コマンドフラグの有効／無効化
        Call Me._CheckAddCommandEnableFlag()
        Call Me._CheckMigrateConditionAddCommandEnableFlag()
        Call Me._CheckPublishQueryCommandEnableFlag()

        AddHandler _
            ConditionChangedEventListener.Instance.DeleteRequested,
            AddressOf Me.DeleteRequestReview
        AddHandler _
            ConditionChangedEventListener.Instance.ChildrenUpdated,
            AddressOf Me._CheckAddCommandEnableFlag
        AddHandler _
            ConditionChangedEventListener.Instance.MigrateConditionUpdated,
            AddressOf Me._CheckMigrateConditionAddCommandEnableFlag
        AddHandler _
            ConditionChangedEventListener.Instance.MigrateConditionDeleteRequested,
            AddressOf Me._MigrateConditionDeleteRequestReview
    End Sub
End Class
