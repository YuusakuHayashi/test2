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
            If String.IsNullOrEmpty(value) Then
                Me._ServerName = DEFAULT_SERVERNAME
            Else
                Me._ServerName = value
            End If
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
            If String.IsNullOrEmpty(value) Then
                Me._DataBaseName = DEFAULT_DATABASENAME
            Else
                Me._DataBaseName = value
            End If
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
            If String.IsNullOrEmpty(value) Then
                Me._DataTableName = DEFAULT_DATATABLENAME
            Else
                Me._DataTableName = value
            End If
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
            If String.IsNullOrEmpty(value) Then
                Me._FieldName = DEFAULT_FIELDNAME
            Else
                Me._FieldName = value
            End If
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
            If String.IsNullOrEmpty(value) Then
                Me._SourceValue = DEFAULT_SOURCEVALUE
            Else
                Me._SourceValue = value
            End If
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
            If String.IsNullOrEmpty(value) Then
                Me._DistinationValue = DEFAULT_DISTINATIONVALUE
            Else
                Me._DistinationValue = value
            End If
            RaisePropertyChanged("DistinationValue")
        End Set
    End Property
    '--------------------------------------------------------------------'


    Sub New(ByRef sql As SqlModel)
        If sql Is Nothing Then
            sql.ServerName = DEFAULT_SERVERNAME
            sql.DataBaseName = DEFAULT_DATABASENAME
            sql.DataTableName = DEFAULT_DATATABLENAME
            sql.FieldName = DEFAULT_FIELDNAME
            sql.SourceValue = DEFAULT_SOURCEVALUE
            sql.DistinationValue = DEFAULT_DISTINATIONVALUE
        End If

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
