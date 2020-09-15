Imports System.Collections.ObjectModel

Public Class MigraterViewModel
    Inherits ProjectBaseViewModel(Of MigraterViewModel)
    ' 条件
    Private _Conditions As ObservableCollection(Of ConditionModel)
    Public Property Conditions As ObservableCollection(Of ConditionModel)
        Get
            Return Me._Conditions
        End Get
        Set(value As ObservableCollection(Of ConditionModel))
            Me._Conditions = value
            RaisePropertyChanged("Conditions")
            Call Me._CheckAddCommandEnabled()
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
            Call Me._CheckMigrateConditionAddEnabled()
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
    Private Sub _CheckAddCommandEnabled()
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
    Private Sub _CheckMigrateConditionAddEnabled()
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
    Private Sub _CheckPublishQueryCommandEnabled()
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
        Dim WherePhrase As String : WherePhrase = _AssembleWherePhrase()

        If Model.ServerName = Me.Model.Server.Name Then
            For Each db In Me.Model.Server.DataBases
                If db.Name = Me.Model.DataBaseName Then
                    For Each dt In db.DataTables
                        dt.IsChecked = False
                        Model.Query = "SELECT COUNT(*) FROM " & dt.Name & " " & WherePhrase
                        Model.DataBaseAccess(AddressOf Model.GetSqlDataSet)
                        If Model.AccessResult Then
                            If Model.QueryResult.Tables(0).Rows(0)(0) > 0 Then
                                dt.IsChecked = True
                                Model.History.AddLine(dt.Name & "は指定条件に一致しました")
                            End If
                        Else
                            If String.IsNullOrEmpty(Model.AccessMessage) Then
                                Model.History.AddLine(Model.AccessMessage)
                            End If
                        End If
                    Next
                End If
            Next
        End If
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


    Private Function _AssembleWherePhrase() As String
        Dim wp As String
        wp = "WHERE "
        wp &= _AssembleWherePhraseBody(Me.Conditions)
        wp &= ";"
        _AssembleWherePhrase = wp
    End Function

    Private Function _AssembleWherePhraseBody(ByVal occm As ObservableCollection(Of ConditionModel)) As String
        Dim wpb As String : wpb = vbNullString
        For Each child In occm
            If child.Logic = "And" Then
                wpb &= " "
                wpb &= child.Logic
            Else
                wpb &= child.Logic
                wpb &= " ("
            End If
            wpb &= " "
            wpb &= child.FieldName
            wpb &= "="
            wpb &= "'" & child.FieldValue & "'"
            If child.Conditions IsNot Nothing Then
                If child.Conditions.Count > 0 Then
                    wpb &= Me._AssembleWherePhraseBody(child.Conditions)
                End If
            End If
            If child.Logic = "And" Then
                wpb &= " "
            Else
                wpb &= ") "
            End If
        Next
        _AssembleWherePhraseBody = wpb
    End Function


    Protected Overrides Sub MyInitializing()
        Me.Conditions = New ObservableCollection(Of ConditionModel) From {
            New ConditionModel
        }

        Me.Migrates = New ObservableCollection(Of MigrateModel) From {
            New MigrateModel
        }
    End Sub


    Protected Overrides Sub ContextModelCheck()
        ViewModel.ContextModel = Me
    End Sub

    Private Sub DeleteRequestAddHandler()
        AddHandler _
            ConditionChangedEventListener.Instance.DeleteRequested,
            AddressOf Me.DeleteRequestReview
    End Sub

    Private Sub ChildrenUpdatedAddHandler()
        AddHandler _
            ConditionChangedEventListener.Instance.ChildrenUpdated,
            AddressOf Me._CheckAddCommandEnabled
    End Sub

    Private Sub MigrateConditionUpdatedAddHandler()
        AddHandler _
            ConditionChangedEventListener.Instance.MigrateConditionUpdated,
            AddressOf Me._CheckMigrateConditionAddEnabled
    End Sub

    Private Sub MigrateConditionDeleteRequestedAddHandler()
        AddHandler _
            ConditionChangedEventListener.Instance.MigrateConditionDeleteRequested,
            AddressOf Me._MigrateConditionDeleteRequestReview
    End Sub

    Sub New(ByRef m As Model,
            ByRef vm As ViewModel)
        Dim ccep(2) As CheckCommandEnabledProxy
        Dim ccep2 As CheckCommandEnabledProxy
        Dim ahp(3) As AddHandlerProxy
        Dim ahp2 As AddHandlerProxy

        ccep(0) = AddressOf Me._CheckAddCommandEnabled
        ccep(1) = AddressOf Me._CheckMigrateConditionAddEnabled
        ccep(2) = AddressOf Me._CheckPublishQueryCommandEnabled

        ahp(0) = AddressOf Me.DeleteRequestAddHandler
        ahp(1) = AddressOf Me.ChildrenUpdatedAddHandler
        ahp(2) = AddressOf Me.MigrateConditionUpdatedAddHandler
        ahp(3) = AddressOf Me.MigrateConditionDeleteRequestedAddHandler

        ccep2 = [Delegate].Combine(ccep)
        ahp2 = [Delegate].Combine(ahp)


        Initializing(m, vm, ccep2, ahp2)
    End Sub
End Class
