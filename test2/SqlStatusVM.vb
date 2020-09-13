Imports System.Reflection
Imports System.ComponentModel

Public Class SqlStatusVM : Inherits BaseViewModel

    Public Const DEFAULT_SERVERNAME = "サーバ名"
    Public Const DEFAULT_DATABASENAME = "データベース名"
    Public Const DEFAULT_DATATABLENAME = "データテーブル名"
    Public Const DEFAULT_FIELDNAME = "フィールド名"
    Public Const DEFAULT_SOURCEVALUE = "複製元フィールド値"
    Public Const DEFAULT_DISTINATIONVALUE = "複製先フィールド値"


    'ViewModelの変更を、Modelに反映
    Private Sub _ModelUpdate()
        If Me.ServerName <> Me.Model.ServerName Then
            Me.Model.ServerName = Me.ServerName
        End If

        If Me.DataBaseName <> Me.Model.DataBaseName Then
            Me.Model.DataBaseName = Me.DataBaseName
        End If

        If Me.DataTableName <> Me.Model.DataTableName Then
            Me.Model.DataTableName = Me.DataTableName
        End If

        If Me.FieldName <> Me.Model.FieldName Then
            Me.Model.FieldName = Me.FieldName
        End If

        If Me.SourceValue <> Me.Model.SourceValue Then
            Me.Model.SourceValue = Me.SourceValue
        End If

        If Me.DistinationValue <> Me.Model.DistinationValue Then
            Me.Model.DistinationValue = Me.DistinationValue
        End If
    End Sub


    'Modelの変更をViewModelに反映
    Private Sub _ViewModelUpdate(ByVal sender As Object,
                                 ByVal e As PropertyChangedEventArgs)
        Dim sm As SqlModel

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

        sm = CType(sender, SqlModel)

        Select Case e.PropertyName
            Case "ServerName"
                Me.ServerName = sm.ServerName
            Case "DataBaseName"
                Me.DataBaseName = sm.DataBaseName
            Case "DataTableName"
                Me.DataTableName = sm.DataTableName
            Case "FieldName"
                Me.FieldName = sm.FieldName
            Case "SourceValue"
                Me.SourceValue = sm.SourceValue
            Case "DistinationValue"
                Me.DistinationValue = sm.DistinationValue
            Case Else
                Exit Sub
        End Select
    End Sub



    'SqlModelへの参照及び変更
    Private _Model As SqlModel
    Public Property Model As SqlModel
        Get
            Return _Model
        End Get
        Set(value As SqlModel)
            _Model = value
        End Set
    End Property


    '-- Server Name -----------------------------------------------------'
    Private _ServerName As String
    Public Property ServerName As String
        Get
            Return _ServerName
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Me._ServerName = DEFAULT_SERVERNAME
            Else
                Me._ServerName = value
            End If
            RaisePropertyChanged("ServerName")
            Call Me._ModelUpdate()
        End Set
    End Property
    '--------------------------------------------------------------------'


    '-- DataBase Name ---------------------------------------------------'
    Private Property _DataBaseName As String
    Public Property DataBaseName As String
        Get
            Return _DataBaseName
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Me._DataBaseName = DEFAULT_DATABASENAME
            Else
                Me._DataBaseName = value
            End If
            RaisePropertyChanged("DataBaseName")
            Call Me._ModelUpdate()
        End Set
    End Property
    '--------------------------------------------------------------------'



    '-- DataTable Name --------------------------------------------------'
    Private Property _DataTableName As String
    Public Property DataTableName As String
        Get
            Return _DataTableName
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Me._DataTableName = DEFAULT_DATATABLENAME
            Else
                Me._DataTableName = value
            End If
            RaisePropertyChanged("DataTableName")
            Call Me._ModelUpdate()
        End Set
    End Property
    '--------------------------------------------------------------------'




    '-- Field Name ------------------------------------------------------'
    Private _FieldName As String
    Public Property FieldName As String
        Get
            Return _FieldName
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Me._FieldName = DEFAULT_FIELDNAME
            Else
                Me._FieldName = value
            End If
            RaisePropertyChanged("FieldName")
            Call Me._ModelUpdate()
        End Set
    End Property
    '--------------------------------------------------------------------'



    '-- Source Value ----------------------------------------------------'
    Private _SourceValue As String
    Public Property SourceValue As String
        Get
            Return _SourceValue
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Me._SourceValue = DEFAULT_SOURCEVALUE
            Else
                Me._SourceValue = value
            End If
            RaisePropertyChanged("SourceValue")
            Call Me._ModelUpdate()
        End Set
    End Property
    '--------------------------------------------------------------------'



    '-- Distination Value -----------------------------------------------'
    Private _DistinationValue As String
    Public Property DistinationValue As String
        Get
            Return _DistinationValue
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Me._DistinationValue = DEFAULT_DISTINATIONVALUE
            Else
                Me._DistinationValue = value
            End If
            RaisePropertyChanged("DistinationValue")
            Call Me._ModelUpdate()
        End Set
    End Property
    '--------------------------------------------------------------------'


    Sub New(ByRef sm As SqlModel)
        Me.Model = sm
        AddHandler Me.Model.PropertyChanged, AddressOf Me._ViewModelUpdate
        'If sql Is Nothing Then
        '    sql.ServerName = DEFAULT_SERVERNAME
        '    sql.DataBaseName = DEFAULT_DATABASENAME
        '    sql.DataTableName = DEFAULT_DATATABLENAME
        '    sql.FieldName = DEFAULT_FIELDNAME
        '    sql.SourceValue = DEFAULT_SOURCEVALUE
        '    sql.DistinationValue = DEFAULT_DISTINATIONVALUE
        'End If

        'If String.IsNullOrEmpty(sql.ServerName) Then
        '    sql.ServerName = DEFAULT_SERVERNAME
        'End If
        'If String.IsNullOrEmpty(sql.DataBaseName) Then
        '    sql.ServerName = DEFAULT_DATABASENAME
        'End If
        'If String.IsNullOrEmpty(sql.DataTableName) Then
        '    sql.ServerName = DEFAULT_DATATABLENAME
        'End If
        'If String.IsNullOrEmpty(sql.FieldName) Then
        '    sql.ServerName = DEFAULT_FIELDNAME
        'End If
        'If String.IsNullOrEmpty(sql.SourceValue) Then
        '    sql.ServerName = DEFAULT_SOURCEVALUE
        'End If
        'If String.IsNullOrEmpty(sql.DistinationValue) Then
        '    sql.ServerName = DEFAULT_DISTINATIONVALUE
        'End If

        'Me.Model = sql

        'Me.ServerName = sql.ServerName
        'Me.DataBaseName = sql.DataBaseName
        'Me.DataTableName = sql.DataTableName
        'Me.FieldName = sql.FieldName
        'Me.SourceValue = sql.SourceValue
        'Me.DistinationValue = sql.DistinationValue

        'AddHandler Me.Model.PropertyChanged, AddressOf Me._ViewModelUpdate
    End Sub
End Class
