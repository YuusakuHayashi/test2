Imports System.ComponentModel
Imports System.Collections.ObjectModel

Public Class ViewModel : Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Protected Sub RaisePropertyChanged(ByVal propertyName As String)
        RaiseEvent PropertyChanged(Me,
                                    New PropertyChangedEventArgs(propertyName))
    End Sub

    Private _Server As ServerModel
    Public Property Server As ServerModel
        Get
            Return _Server
        End Get
        Set(value As ServerModel)
            _Server = value
        End Set
    End Property

    Private _ServerName As String
    Public Property ServerName As String
        Get
            Return _ServerName
        End Get
        Set(value As String)
            _ServerName = value
        End Set
    End Property

    Private _DataBase As List(Of DataBase)
    Public Property DataBase As List(Of DataBase)
        Get
            Return _DataBase
        End Get
        Set(value As List(Of DataBase))
            _DataBase = value
        End Set
    End Property

    Private _DataBaseName As String
    Public Property DataBaseName As String
        Get
            Return _DataBaseName
        End Get
        Set(value As String)
            _DataBaseName = value
        End Set
    End Property

    Private _DataTables As List(Of DataTable)
    Public Property DataTables As List(Of DataTable)
        Get
            Return _DataTables
        End Get
        Set(value As List(Of DataTable))
            _DataTables = value
        End Set
    End Property

    Private _DataTable As DataTable
    Public Property DataTable As DataTable
        Get
            Return _DataTable
        End Get
        Set(value As DataTable)
            _DataTable = value
        End Set
    End Property

    Private _DataTableName As String
    Public Property DataTableName As String
        Get
            Return _DataTableName
        End Get
        Set(value As String)
            _DataTableName = value
        End Set
    End Property

    Private _Persons As ObservableCollection(Of Person)
    Public Property Persons As ObservableCollection(Of Person)
        Get
            Return _Persons
        End Get
        Set(value As ObservableCollection(Of Person))
            _Persons = value
        End Set
    End Property

    Private _Child As ObservableCollection(Of Person)
    Public Property Child As ObservableCollection(Of Person)
        Get
            Return _Child
        End Get
        Set(value As ObservableCollection(Of Person))
            _Child = value
        End Set
    End Property


    'Private _Name As String
    'Public Property Name As String
    '    Get
    '        Return _Name
    '    End Get
    '    Set(value As String)
    '        _Name = value
    '    End Set
    'End Property


    Public Sub New()
        Me.Persons = New ObservableCollection(Of Person)
        Me.Persons.Add(New Person With {
            .Name = "hoge",
            .Child = New ObservableCollection(Of Person) From {
                New Person With {.Name = "fuga"},
                New Person With {.Name = "piyo"}
            }
        })
        'Me.DataTables = New List(Of DataTables)
        'Me.DataTables.Add(New DataTables With {
        '    .DataTableName = "hoge",
        '    .DataTable = New List(Of DataTables)
        '})
        'Me.DataTable.DataTable.Add(New DataTables With {.DataTableName = "fuga"})
        'Dim m As Model = New Model
        'm.Name = "hogehoge"
        'Dim s As ServerModel = New ServerModel _
        '    With {.ServerName = "hoge", .DataBases = New List(Of DataBaseModel)}
        'Dim db As DataBaseModel = New DataBaseModel _
        '    With {.DataBaseName = "fuga", .DataTables = New List(Of DataTableModel)}
        'Dim dt As DataTableModel = New DataTableModel _
        '    With {.DataTableName = "piyo"}


        'db.DataTables.Add(dt)
        's.DataBases.Add(db)
        ''Me.ServerName = s.ServerName
        'Me.DataTable = New List(Of DataTableModel)({dt, dt, dt})
    End Sub
End Class
