Imports System.ComponentModel
Imports System.Collections.ObjectModel

Delegate Sub MySub()

Public Class ViewModel : Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Protected Sub RaisePropertyChanged(ByVal propertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(propertyName)
        )
    End Sub

    Private _ServerVM As List(Of ServerModel)
    Public Property ServerVM As List(Of ServerModel)
        Get
            Return _ServerVM
        End Get
        Set(value As List(Of ServerModel))
            _ServerVM = value
        End Set
    End Property

    Private __VM As Model
    Private Property _VM As Model
        Get
            Return __VM
        End Get
        Set(value As Model)
            __VM = value
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

    Private Property _DataBaseName As String
    Public Property DataBaseName As String
        Get
            Return _DataBaseName
        End Get
        Set(value As String)
            _DataBaseName = value
        End Set
    End Property

    Private Property _DataTableName As String
    Public Property DataTableName As String
        Get
            Return _DataTableName
        End Get
        Set(value As String)
            _DataTableName = value
        End Set
    End Property

    Private Property _FieldName As String
    Public Property FieldName As String
        Get
            Return _FieldName
        End Get
        Set(value As String)
            _FieldName = value
        End Set
    End Property

    Private Property _SourceValue As String
    Public Property SourceValue As String
        Get
            Return _SourceValue
        End Get
        Set(value As String)
            _SourceValue = value
        End Set
    End Property

    Private Property _DistinationValue As String
    Public Property DistinationValue As String
        Get
            Return _DistinationValue
        End Get
        Set(value As String)
            _DistinationValue = value
        End Set
    End Property


    Private _IniFile As String
    Public Property IniFile As String
        Get
            Return _IniFile
        End Get
        Set(value As String)
            _IniFile = value
        End Set
    End Property

    Private _ConfigFile As String
    Public Property ConfigFile As String
        Get
            Return _ConfigFile
        End Get
        Set(value As String)
            _ConfigFile = value
        End Set
    End Property

    Private _MyDirectory As String
    Public Property MyDirectory As String
        Get
            Return _MyDirectory
        End Get
        Set(value As String)
            _MyDirectory = value
        End Set
    End Property

    Private Property _ServerSelection As List(Of ServerModel)
    Public Property ServerSelection As List(Of ServerModel)
        Get
            Return _ServerSelection
        End Get
        Set(value As List(Of ServerModel))
            _ServerSelection = value
        End Set
    End Property

    Private Property _DataBaseSelection As List(Of String)
    Public Property DataBaseSelection As List(Of String)
        Get
            Return _DataBaseSelection
        End Get
        Set(value As List(Of String))
            _DataBaseSelection = value
        End Set
    End Property

    Private Property _DataTableSelection As List(Of String)
    Public Property DataTableSelection As List(Of String)
        Get
            Return _DataTableSelection
        End Get
        Set(value As List(Of String))
            _DataTableSelection = value
        End Set
    End Property

    Private Sub SaveCommandExecute(ByVal p As Object)
    End Sub

    Private _SaveCommand As ICommand
    Public ReadOnly Property SaveCommand As ICommand
        Get
            If _SaveCommand Is Nothing Then
                _SaveCommand = New DelegateCommand With {
                    .ExecuteDelegater = AddressOf ,
                    .CanExecuteDelegater = AddressOf
                }
            End If
        End Get
    End Property


    Sub New()
        Me._VM = New Model
        Me.ServerVM = New List(Of ServerModel)

        Me._VM.Servers = Me._VM.ServersLoad()

        Me.ServerVM = Me._VM.Servers
        Me.ServerName = Me._VM.Servers(0).ServerName
        Me.DataBaseName = Me._VM.Servers(0).DataBaseName
        Me.DataTableName = Me._VM.Servers(0).DataTableName
        'Me._VM.Servers = New List(Of ServerModel) From {
        '    New ServerModel With {
        '        .RealName = "AAA", .Child = New List(Of ServerModel) From {
        '            New ServerModel With {
        '                .RealName = "ABA"
        '            },
        '            New ServerModel With {
        '                .RealName = "ACA"
        '            }
        '        }
        '    },
        '    New ServerModel With {
        '        .RealName = "BAA", .Child = New List(Of ServerModel) From {
        '            New ServerModel With {
        '                .RealName = "BBA"
        '            },
        '            New ServerModel With {
        '                .RealName = "BCA"
        '            }
        '        }
        '    }
        '}
        'Me.ServerVM = Me._VM.Servers
        'Me._VM.ModelSave()
        'Dim s As ServerModel
        'Me._VM = New Model
        's = Me._VM.ServerLoad
        'Me.ServerVM = New List(Of ServerModel)
        'Me.ServerVM.Add(s)
    End Sub
End Class

Public Class DelegateCommand : Implements ICommand

    Public Property CanExecuteDelegater As Func(Of Object, Boolean)
    Public Property ExecuteDelegater As Action(Of String)

    Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

    Public Sub RaiseCanExecuteChanged()
        RaiseEvent CanExecuteChanged(Me, Nothing)
    End Sub

    Public Sub Execute(parameter As Object) Implements ICommand.Execute
        Dim d = ExecuteDelegater
        If d <> Nothing Then
            d(parameter)
        End If
    End Sub

    Public Function CanExecute(ByVal parameter As Object) As Boolean Implements ICommand.CanExecute
        Dim d = CanExecuteDelegater
        Return IIf(d = Nothing, True, d(parameter))
    End Function
End Class
