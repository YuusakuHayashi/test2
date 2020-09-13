'Imports System.ComponentModel
'Imports System.Collections.ObjectModel


'Public Class AccessButtonVM
'    Inherits BaseViewModel

'    Private __Model As SqlModel
'    Private Property _Model As SqlModel
'        Get
'            Return __Model
'        End Get
'        Set(value As SqlModel)
'            __Model = value
'        End Set
'    End Property


'    'SqlModel変更の反映, プロパティの意識無し
'    Private Sub _ModelUpdate(ByVal sender As Object,
'                              ByVal e As PropertyChangedEventArgs)
'        Dim sm As SqlModel
'        sm = CType(sender, SqlModel)
'        _Model = sm

'        'イベントハンドラに追加していくよりも
'        'ここにViewModel変更時のトリガーを追加していったほうがいい・・・？
'        Call Me._AccessEnableCheck()
'    End Sub

'    Private Sub _AccessEnableCheck()
'        If Me._Model.ServerName <> SqlStatusVM.DEFAULT_SERVERNAME _
'            And (Not String.IsNullOrEmpty(Me._Model.ServerName)) Then
'            If Me._Model.DataBaseName <> SqlStatusVM.DEFAULT_DATABASENAME _
'                And (Not String.IsNullOrEmpty(Me._Model.DataBaseName)) Then
'                If Me._Model.DataTableName <> SqlStatusVM.DEFAULT_DATATABLENAME _
'                    And (Not String.IsNullOrEmpty(Me._Model.DataTableName)) Then
'                    If Me._Model.FieldName <> SqlStatusVM.DEFAULT_FIELDNAME _
'                        And (Not String.IsNullOrEmpty(Me._Model.FieldName)) Then
'                        If Me._Model.SourceValue <> SqlStatusVM.DEFAULT_SOURCEVALUE _
'                        And (Not String.IsNullOrEmpty(Me._Model.SourceValue)) Then
'                            If Me._Model.DistinationValue <> SqlStatusVM.DEFAULT_DISTINATIONVALUE _
'                                And (Not String.IsNullOrEmpty(Me._Model.DistinationValue)) Then
'                                Me.AccessEnableFlag = True
'                                Exit Sub
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If

'        Me.AccessEnableFlag = False
'    End Sub


'    Private _AccessCommand As ICommand
'    Public ReadOnly Property AccessCommand As ICommand
'        Get
'            If _AccessCommand Is Nothing Then
'                _AccessCommand = New DelegateCommand With {
'                    .ExecuteHandler = AddressOf _AccessCommandExecute,
'                    .CanExecuteHandler = AddressOf _AceessCommandCanExecute
'                }
'                Return _AccessCommand
'            Else
'                Return _AccessCommand
'            End If
'        End Get
'    End Property

'    Private _AccessEnableFlag As Boolean
'    Public Property AccessEnableFlag As Boolean
'        Get
'            Return _AccessEnableFlag
'        End Get
'        Set(value As Boolean)
'            _AccessEnableFlag = value
'            RaisePropertyChanged("AccessEnableFlag")
'            CType(AccessCommand, DelegateCommand).RaiseCanExecuteChanged()
'        End Set
'    End Property

'    Private Sub _AccessCommandExecute(ByVal parameter As Object)

'        'テスト用
'        'Call Me._Model.AccessTestProxy(True)

'        '本番
'        'Call Me._Model.AccessTest()
'    End Sub

'    Private Function _AceessCommandCanExecute(ByVal parameter As Object) As Boolean
'        Return AccessEnableFlag
'    End Function

'    Sub New(ByRef sm As SqlModel, ByRef ssvm As SqlStatusVM)
'        _Model = sm
'        '_VM = ssvm

'        Call Me._AccessEnableCheck()

'        'AddHandler ssvm.PropertyChanged, AddressOf _VMUpdate
'        AddHandler sm.PropertyChanged, AddressOf _ModelUpdate
'    End Sub
'End Class
