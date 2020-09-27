Imports System.Collections.ObjectModel
Imports System.Data.SqlClient

Public Class MigraterViewModel
    Inherits ProjectBaseViewModel(Of MigraterViewModel)

    'セーブ対象クラス
    Public Class SaveData
        Public Conditions As List(Of SaveConditionModel)
        Public Migrates As List(Of SaveMigrateModel)
        Public TargetTables As String
    End Class

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


    Private Const CHECKED_TABLES As String = "チェックされたテーブルを対象"
    Private Const ALL_TABLES As String = "全てのテーブルが対象"

    Public ReadOnly Property TargetTablesList As List(Of String)
        Get
            Return New List(Of String) From {
                CHECKED_TABLES,
                ALL_TABLES
            }
        End Get
    End Property

    Private _TargetTables As String
    Public Property TargetTables As String
        Get
            Return Me._TargetTables
        End Get
        Set(value As String)
            Me._TargetTables = value
            RaisePropertyChanged("TargetTables")
            Call Me._CheckUpdateCommandEnabled()
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


    '--- ＡＤＤ ------------------------------------------------------------------------'
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

    ' コマンド有効／無効化(ＡＤＤ)
    Private Function _AddCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._AddCommandEnableFlag
    End Function

    '-----------------------------------------------------------------------------------'


    '--- 移行後ＡＤＤ ------------------------------------------------------------------'
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

    ' コマンド有効／無効化(移行後ＡＤＤ)
    Private Function _MigrateConditionAddCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._MigrateConditionAddCommandEnableFlag
    End Function
    '-----------------------------------------------------------------------------------'


    '--- ＣＫクエリ --------------------------------------------------------------------'
    ' コマンドプロパティ(ＣＫクエリ)
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

    ' コマンド有効／無効判定(ＣＫクエリ)
    Private Sub _CheckPublishQueryCommandEnabled()
        Dim b As Boolean : b = True
        Me._PublishQueryCommandEnableFlag = b
    End Sub

    'コマンド実行可否のフラグ(ＣＫクエリ)
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

    ' コマンド実行(ＣＫクエリ)
    Private Sub _PublishQueryCommandExecute(ByVal parameter As Object)
        Dim WherePhrase As String : WherePhrase = _AssembleWherePhrase()
        Dim query As String
        Dim mysql As New MySql

        Try
            mysql.ConnectionString = Model.ConnectionString
            Call mysql.ConnectionStart()
            For Each db In Me.Model.Server.DataBases
                If db.Name = Me.Model.DataBaseName Then
                    For Each dt In db.DataTables
                        If dt.IsChecked Then
                            dt.IsChecked = False
                        End If
                        query = "SELECT COUNT(*) FROM " & dt.Name & " " & WherePhrase
                        Try
                            Call mysql.GetSqlData(query)
                            If mysql.Result.Tables(0).Rows(0)(0) > 0 Then
                                dt.IsChecked = True
                                Model.History.AddLine(dt.Name & "は指定条件に一致しました")
                            Else
                                Model.History.AddLine(dt.Name & "は指定条件に一致しませんでした")
                            End If
                        Catch ex As Exception
                            If mysql.ErrorResult IsNot Nothing Then
                                Model.History.AddLine(ex.Message)
                            End If
                        Finally
                        End Try
                    Next
                End If
            Next
            Call mysql.ConnectionCommit()
        Catch ex As Exception
            Model.History.AddLine(ex.Message)
            Call mysql.ConnectionFailed()
        Finally
            Call mysql.ConnectionClose()
            Call mysql.DeleteTable(Model.TestTableName)
        End Try
    End Sub

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

    ' コマンド有効／無効化(ＣＫクエリ)
    Private Function _PublishQueryCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._PublishQueryCommandEnableFlag
    End Function
    '-----------------------------------------------------------------------------------'

    '-----------------------------------------------------------------------------------'
    ' コマンドプロパティ(更新クエリ)
    Private _UpdateCommand As ICommand
    Public ReadOnly Property UpdateCommand As ICommand
        Get
            If Me._UpdateCommand Is Nothing Then
                Me._UpdateCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _UpdateCommandExecute,
                    .CanExecuteHandler = AddressOf _UpdateCommandCanExecute
                }
                Return Me._UpdateCommand
            Else
                Return Me._UpdateCommand
            End If
        End Get
    End Property

    ' コマンド有効／無効判定(更新クエリ)
    Private Sub _CheckUpdateCommandEnabled()
        Dim b As Boolean : b = False
        Select Case TargetTables
            Case CHECKED_TABLES
                For Each db In Me.Model.Server.DataBases
                    If db.Name = Me.Model.DataBaseName Then
                        For Each dt In db.DataTables
                            If dt.IsChecked Then
                                b = True
                                Exit For
                            End If
                        Next
                        If b Then
                            Exit For
                        End If
                    End If
                Next
            Case ALL_TABLES
                b = True
            Case Else
                'b = False
        End Select
        _UpdateCommandEnableFlag = b
    End Sub

    'コマンド実行可否のフラグ(更新クエリ)
    Private __UpdateCommandEnableFlag As Boolean
    Public Property _UpdateCommandEnableFlag As Boolean
        Get
            Return Me.__UpdateCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__UpdateCommandEnableFlag = value
            RaisePropertyChanged("_UpdateCommandEnableFlag")
            CType(UpdateCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    ' コマンド実行(更新クエリ)
    Private Sub _UpdateCommandExecute(ByVal parameter As Object)
        Dim WherePhrase As String : WherePhrase = _AssembleWherePhrase()
        Dim query As String
        Dim mysql As New MySql

        Try
            mysql.ConnectionString = Model.ConnectionString
            Call mysql.ConnectionStart()
            For Each db In Me.Model.Server.DataBases
                If db.Name = Me.Model.DataBaseName Then
                    For Each dt In db.DataTables
                        If TargetTables = CHECKED_TABLES Then
                            If Not dt.IsChecked Then
                                Continue For
                            End If
                        End If
                        If dt.IsChecked Then
                            ' テストテーブルの削除
                            query = $"  IF OBJECT_ID('{Model.TestTableName}', 'U') IS NOT NULL"
                            query &= $"    BEGIN DROP TABLE {Model.TestTableName}"
                            query &= $" END"
                            Call mysql.ExecuteSql(query)

                            ' 空テーブル作成
                            query = $"  SELECT * INTO {Model.TestTableName}"
                            query &= $"    FROM {dt.Name}"
                            query &= $"       WHERE 1 <> 1"
                            Call mysql.ExecuteSql(query)

                            ' テストテーブル複製
                            query = $"  INSERT INTO {Model.TestTableName}"
                            query &= $"    SELECT * FROM {dt.Name} " & WherePhrase
                            Call mysql.ExecuteSql(query)

                            ' テストテーブル更新
                            query = $"  UPDATE {Model.TestTableName}"
                            query &= $"    SET"
                            For Each m In Migrates
                                query &= $"   {m.FieldName} = '{m.FieldValue}'"
                                If m IsNot Migrates.Last() Then
                                    query &= ","
                                End If
                            Next
                            query &= $" " & WherePhrase
                            Try
                                Call mysql.ExecuteSql(query)
                            Catch ex As Exception
                            End Try

                            ' テストテーブルから挿入
                            Try
                                query = $"  INSERT INTO {dt.Name}"
                                query &= $"    SELECT * FROM {Model.TestTableName}"
                                Call mysql.ExecuteSql(query)
                            Catch ex As Exception
                                If mysql.ErrorResult.Number = 2627 Then
                                    Model.History.AddLine(ex.Message)
                                End If
                            Finally
                            End Try
                        End If
                    Next
                End If
            Next
            mysql.ConnectionCommit()
        Catch ex As Exception
            Model.History.AddLine(ex.Message)
            Call mysql.ConnectionFailed()
        Finally
            Call mysql.ConnectionClose()
            Call mysql.DeleteTable(Model.TestTableName)
        End Try
    End Sub


    ' コマンド有効／無効化(更新クエリ)
    Private Function _UpdateCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._UpdateCommandEnableFlag
    End Function
    '-----------------------------------------------------------------------------------'


    '-----------------------------------------------------------------------------------'
    ' コマンドプロパティ(ビューセーブ)
    Private _ViewSaveCommand As ICommand
    Public ReadOnly Property ViewSaveCommand As ICommand
        Get
            If Me._ViewSaveCommand Is Nothing Then
                Me._ViewSaveCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ViewSaveCommandExecute,
                    .CanExecuteHandler = AddressOf _ViewSaveCommandCanExecute
                }
                Return Me._ViewSaveCommand
            Else
                Return Me._ViewSaveCommand
            End If
        End Get
    End Property

    ' コマンド有効／無効判定(ビューセーブ)
    Private Sub _CheckViewSaveCommandEnabled()
        Dim b As Boolean : b = True
        _ViewSaveCommandEnableFlag = b
    End Sub

    'コマンド実行可否のフラグ(ビューセーブ)
    Private __ViewSaveCommandEnableFlag As Boolean
    Public Property _ViewSaveCommandEnableFlag As Boolean
        Get
            Return Me.__ViewSaveCommandEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ViewSaveCommandEnableFlag = value
            RaisePropertyChanged("_ViewSaveCommandEnableFlag")
            CType(ViewSaveCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    ' コマンド実行(ビューセーブ)
    Private Sub _ViewSaveCommandExecute(ByVal parameter As Object)
        Dim sd As New SaveData
        sd.Conditions = ConvertToSaveConditions(Me.Conditions)
        sd.Migrates = ConvertToSaveMigrates(Me.Migrates)
        Call ModelSave(Of SaveData)("Test.json", sd)
    End Sub


    ' コマンド有効／無効化(ビューセーブ)
    Private Function _ViewSaveCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ViewSaveCommandEnableFlag
    End Function
    '-----------------------------------------------------------------------------------'

    ' ＣｏｎｄｉｔｉｏｎモデルはＩＣｏｍｍａｎｄプロパティを持ってしまっているため、除いたモデルに変換する
    Private Function ConvertToSaveConditions(ByVal lcm As ObservableCollection(Of ConditionModel)) As List(Of SaveConditionModel)
        Dim lscm As New List(Of SaveConditionModel)
        Dim recursive As Func(Of ConditionModel, SaveConditionModel)

        recursive = Function(ByVal cm As ConditionModel) As SaveConditionModel
                        Dim scm As New SaveConditionModel

                        scm.Logic = cm.Logic
                        scm.FieldName = cm.FieldName
                        scm.FieldValue = cm.FieldValue
                        scm.Conditions = New List(Of SaveConditionModel)

                        If cm.Conditions IsNot Nothing Then
                            For Each bro In cm.Conditions
                                scm.Conditions.Add(recursive(bro))
                            Next
                        End If

                        Return scm
                    End Function

        If lcm IsNot Nothing Then
            For Each bro In lcm
                lscm.Add(recursive(bro))
            Next
        End If

        ConvertToSaveConditions = lscm
    End Function

    ' ＭｉｇｒａｔｅｓモデルはＩＣｏｍｍａｎｄプロパティを持ってしまっているため、除いたモデルに変換する
    Private Function ConvertToSaveMigrates(ByVal lmm As ObservableCollection(Of MigrateModel)) As List(Of SaveMigrateModel)
        Dim lsmm As New List(Of SaveMigrateModel)
        Dim recursive As Func(Of MigrateModel, SaveMigrateModel)

        recursive = Function(ByVal mm As MigrateModel) As SaveMigrateModel
                        Dim smm As New SaveMigrateModel

                        smm.Logic = mm.Logic
                        smm.FieldName = mm.FieldName
                        smm.FieldValue = mm.FieldValue

                        Return smm
                    End Function

        If lmm IsNot Nothing Then
            For Each bro In lmm
                lsmm.Add(recursive(bro))
            Next
        End If

        ConvertToSaveMigrates = lsmm
    End Function

    Private Sub _DataTableCheckChangedReview(ByVal sender As Object, ByVal e As EventArgs)
        Call Me._CheckUpdateCommandEnabled()
    End Sub

    Protected Overrides Sub MyInitializing()
        Me.Conditions = New ObservableCollection(Of ConditionModel) From {
            New ConditionModel
        }
        Me.Migrates = New ObservableCollection(Of MigrateModel) From {
            New MigrateModel
        }
    End Sub


    Protected Overrides Sub ContextModelCheck()
        ViewModel.SetContext(ViewModel.MAIN_VIEW, Me.GetType.Name, Me)
    End Sub

    Private Sub DeleteRequestAddHandler()
        AddHandler _
            MyEventListener.Instance.DeleteRequested,
            AddressOf Me.DeleteRequestReview
    End Sub

    Private Sub ChildrenUpdatedAddHandler()
        AddHandler _
            MyEventListener.Instance.ChildrenUpdated,
            AddressOf Me._CheckAddCommandEnabled
    End Sub

    Private Sub MigrateConditionUpdatedAddHandler()
        AddHandler _
            MyEventListener.Instance.MigrateConditionUpdated,
            AddressOf Me._CheckMigrateConditionAddEnabled
    End Sub

    Private Sub MigrateConditionDeleteRequestedAddHandler()
        AddHandler _
            MyEventListener.Instance.MigrateConditionDeleteRequested,
            AddressOf Me._MigrateConditionDeleteRequestReview
    End Sub

    Private Sub _DataTableCheckChangedAddHandler()
        AddHandler _
            MyEventListener.Instance.DataTableCheckChanged,
            AddressOf Me._DataTableCheckChangedReview
    End Sub

    Sub New(ByRef m As Model,
            ByRef vm As ViewModel)

        Dim ccep(4) As CheckCommandEnabledProxy
        Dim ccep2 As CheckCommandEnabledProxy
        Dim ahp(4) As AddHandlerProxy
        Dim ahp2 As AddHandlerProxy

        ccep(0) = AddressOf Me._CheckAddCommandEnabled
        ccep(1) = AddressOf Me._CheckMigrateConditionAddEnabled
        ccep(2) = AddressOf Me._CheckPublishQueryCommandEnabled
        ccep(3) = AddressOf Me._CheckUpdateCommandEnabled
        ccep(4) = AddressOf Me._CheckViewSaveCommandEnabled

        ahp(0) = AddressOf Me.DeleteRequestAddHandler
        ahp(1) = AddressOf Me.ChildrenUpdatedAddHandler
        ahp(2) = AddressOf Me.MigrateConditionUpdatedAddHandler
        ahp(3) = AddressOf Me.MigrateConditionDeleteRequestedAddHandler
        ahp(4) = AddressOf Me._DataTableCheckChangedAddHandler

        ccep2 = [Delegate].Combine(ccep)
        ahp2 = [Delegate].Combine(ahp)


        Initializing(m, vm, ccep2, ahp2)
    End Sub
End Class
