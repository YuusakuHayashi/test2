Imports System.ComponentModel
Imports System.Collections.ObjectModel


Public Class AccessButtonVM
    Inherits ViewModel

    Private __Model As SqlModel
    Private Property _Model As SqlModel
        Get
            Return __Model
        End Get
        Set(value As SqlModel)
            __Model = value
        End Set
    End Property

    Private __VM As SqlStatusVM
    Private Property _VM As SqlStatusVM
        Get
            Return __VM
        End Get
        Set(value As SqlStatusVM)
            __VM = value
        End Set
    End Property


    'SqlModel変更の反映, プロパティの意識無し
    Private Sub _ModelUpdate(ByVal sender As Object,
                              ByVal e As PropertyChangedEventArgs)
        Dim sm As SqlModel
        sm = CType(sender, SqlModel)
        _Model = sm
    End Sub

    'SqlStatusVM変更の反映, プロパティの意識無し
    Private Sub _VMUpdate(ByVal sender As Object,
                              ByVal e As PropertyChangedEventArgs)
        Dim ssvm As SqlStatusVM

        ssvm = CType(sender, SqlStatusVM)
        _VM = ssvm
    End Sub



    Private Sub _AccessEnableCheck(ByVal sender As Object,
                                  ByVal e As PropertyChangedEventArgs)
        Dim ssvm As SqlStatusVM

        Select Case e.PropertyName
            Case "ServerName"
            Case "DataBaseName"
            Case "DataTableName"
            Case "FieldName"
            Case "SourceValue"
            Case "DistinationValue"
            Case Else
                Exit Sub
        End Select

        ssvm = CType(sender, SqlStatusVM)
        If ssvm.ServerName <> SqlStatusVM.DEFAULT_SERVERNAME _
            And (Not String.IsNullOrEmpty(ssvm.ServerName)) Then
            If ssvm.DataBaseName <> SqlStatusVM.DEFAULT_DATABASENAME _
                And (Not String.IsNullOrEmpty(ssvm.DataBaseName)) Then
                If ssvm.DataTableName <> SqlStatusVM.DEFAULT_DATATABLENAME _
                    And (Not String.IsNullOrEmpty(ssvm.DataTableName)) Then
                    If ssvm.FieldName <> SqlStatusVM.DEFAULT_FIELDNAME _
                        And (Not String.IsNullOrEmpty(ssvm.FieldName)) Then
                        If ssvm.SourceValue <> SqlStatusVM.DEFAULT_SOURCEVALUE _
                        And (Not String.IsNullOrEmpty(ssvm.SourceValue)) Then
                            If ssvm.DistinationValue <> SqlStatusVM.DEFAULT_DISTINATIONVALUE _
                                And (Not String.IsNullOrEmpty(ssvm.DistinationValue)) Then
                                Me.AccessEnableFlag = True
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If


        Me.AccessEnableFlag = False

    End Sub


    Private _AccessCommand As ICommand
    Public ReadOnly Property AccessCommand As ICommand
        Get
            If _AccessCommand Is Nothing Then
                _AccessCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _AccessCommandExecute,
                    .CanExecuteHandler = AddressOf _AceessCommandCanExecute
                }
                Return _AccessCommand
            Else
                Return _AccessCommand
            End If
        End Get
    End Property

    Private _AccessEnableFlag As Boolean
    Public Property AccessEnableFlag As Boolean
        Get
            Return _AccessEnableFlag
        End Get
        Set(value As Boolean)
            _AccessEnableFlag = value
            RaisePropertyChanged("AccessEnableFlag")
            CType(AccessCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _AccessCommandExecute(ByVal parameter As Object)
        Call _Model.AccessTest()
    End Sub

    Private Function _AceessCommandCanExecute(ByVal parameter As Object) As Boolean
        Return AccessEnableFlag
    End Function

    Sub New(ByRef sm As SqlModel, ByRef ssvm As SqlStatusVM)

        _Model = sm
        _VM = ssvm

        Me.AccessEnableFlag = False

        AddHandler ssvm.PropertyChanged, AddressOf _VMUpdate
        AddHandler ssvm.PropertyChanged, AddressOf _AccessEnableCheck

        AddHandler sm.PropertyChanged, AddressOf _ModelUpdate
    End Sub
End Class
