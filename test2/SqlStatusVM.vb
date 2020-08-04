Imports System.Reflection
Imports System.ComponentModel

Public Class SqlStatusVM : Inherits ViewModel

    Public Const DEFAULT_SERVERNAME = "サーバ名"
    Public Const DEFAULT_DATABASENAME = "データベース名"
    Public Const DEFAULT_DATATABLENAME = "データテーブル名"
    Public Const DEFAULT_FIELDNAME = "フィールド名"
    Public Const DEFAULT_SOURCEVALUE = "複製元フィールド値"
    Public Const DEFAULT_DISTINATIONVALUE = "複製先フィールド値"


    'モデルをUpdate
    Private Sub _ModelUpdate(ByVal sender As Object,
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
        Select Case e.PropertyName
            Case "ServerName"
                Me._Model.ServerName = ssvm.ServerName
            Case "DataBaseName"
                Me._Model.DataBaseName = ssvm.DataBaseName
            Case "DataTableName"
                Me._Model.DataTableName = ssvm.DataTableName
            Case "FieldName"
                Me._Model.FieldName = ssvm.FieldName
            Case "SourceValue"
                Me._Model.SourceValue = ssvm.SourceValue
            Case "DistinationValue"
                Me._Model.DistinationValue = ssvm.DistinationValue
            Case Else
                Exit Sub
        End Select
    End Sub

    Private __Model As SqlModel
    Private Property _Model As SqlModel
        Get
            Return __Model
        End Get
        Set(value As SqlModel)
            __Model = value
        End Set
    End Property


    '-- Server Name -----------------------------------------------------'
    Private _ServerName As String
    Public Property ServerName As String
        Get
            Return _ServerName
        End Get
        Set(value As String)
            _ServerName = value
            RaisePropertyChanged("ServerName")
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
            _DataBaseName = value
            RaisePropertyChanged("DataBaseName")
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
            _DataTableName = value
            RaisePropertyChanged("DataTableName")
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
            _FieldName = value
            RaisePropertyChanged("FieldName")
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
            _SourceValue = value
            RaisePropertyChanged("SourceValue")
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
            _DistinationValue = value
            RaisePropertyChanged("DistinationValue")
        End Set
    End Property
    '--------------------------------------------------------------------'


    '初期化
    Private Function _SqlModelInitialize() As SqlModel
        Dim sm As New SqlModel
        sm.ServerName = DEFAULT_SERVERNAME
        sm.DataBaseName = DEFAULT_DATABASENAME
        sm.DataTableName = DEFAULT_DATATABLENAME
        sm.FieldName = DEFAULT_FIELDNAME
        sm.SourceValue = DEFAULT_SOURCEVALUE
        sm.DistinationValue = DEFAULT_DISTINATIONVALUE
        _SqlModelInitialize = sm
    End Function


    Sub New(ByRef sql As SqlModel)
        If sql Is Nothing Then
            sql = Me._SqlModelInitialize
        End If

        'Dim mi As MemberInfo
        'Dim mis As MemberInfo()
        'mis = GetType(SqlModel).GetMembers()
        'For Each mi In mis
        '    If mi.MemberType = MemberTypes.Property Then
        '        If String.IsNullOrEmpty(GetType(SqlModel).GetProperty(mi.Name).GetValue(sql)) Then

        '        End If
        '    End If
        'Next

        If String.IsNullOrEmpty(sql.ServerName) Then
            sql.ServerName = DEFAULT_SERVERNAME
        End If
        If String.IsNullOrEmpty(sql.DataBaseName) Then
            sql.ServerName = DEFAULT_DATABASENAME
        End If
        If String.IsNullOrEmpty(sql.DataTableName) Then
            sql.ServerName = DEFAULT_DATATABLENAME
        End If
        If String.IsNullOrEmpty(sql.FieldName) Then
            sql.ServerName = DEFAULT_FIELDNAME
        End If
        If String.IsNullOrEmpty(sql.SourceValue) Then
            sql.ServerName = DEFAULT_SOURCEVALUE
        End If
        If String.IsNullOrEmpty(sql.DistinationValue) Then
            sql.ServerName = DEFAULT_DISTINATIONVALUE
        End If

        Me._Model = sql

        Me.ServerName = sql.ServerName
        Me.DataBaseName = sql.DataBaseName
        Me.DataTableName = sql.DataTableName
        Me.FieldName = sql.FieldName
        Me.SourceValue = sql.SourceValue
        Me.DistinationValue = sql.DistinationValue

        AddHandler Me.PropertyChanged, AddressOf Me._ModelUpdate
    End Sub
End Class
