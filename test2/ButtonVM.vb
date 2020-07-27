Imports System.ComponentModel
Imports System.Collections.ObjectModel


Public Class ButtonVM
    Inherits ViewModel

    Private _AccessEnableFlag As Boolean
    Public Property AccessEnableFlag As Boolean
        Get
            Return _AccessEnableFlag
        End Get
        Set(value As Boolean)
            _AccessEnableFlag = value
        End Set
    End Property


    Sub New()
        Dim abvm As New AccessButtonVM
    End Sub
End Class


Public Class AccessButtonVM
    Inherits ViewModel

    Private _ConfigFileName As String

    Private ReadOnly Property _EnableFlag As Boolean
        Get
            _ConfigFileName = (New ConfigFileVM).ConfigFileName
            If _ConfigFileName <> vbNullString Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Sub New()
    End Sub

    'Private _ServerName As String
    'Public Property ServerName As String
    '    Get
    '        Return _ServerName
    '    End Get
    '    Set(value As String)
    '        _ServerName = value
    '    End Set
    'End Property

    'Private Property _DataBaseName As String
    'Public Property DataBaseName As String
    '    Get
    '        Return _DataBaseName
    '    End Get
    '    Set(value As String)
    '        _DataBaseName = value
    '    End Set
    'End Property

    'Private Property _DataTableName As String
    'Public Property DataTableName As String
    '    Get
    '        Return _DataTableName
    '    End Get
    '    Set(value As String)
    '        _DataTableName = value
    '    End Set
    'End Property

    'Private Property _FieldName As String
    'Public Property FieldName As String
    '    Get
    '        Return _FieldName
    '    End Get
    '    Set(value As String)
    '        _FieldName = value
    '    End Set
    'End Property

    'Private Property _SourceValue As String
    'Public Property SourceValue As String
    '    Get
    '        Return _SourceValue
    '    End Get
    '    Set(value As String)
    '        _SourceValue = value
    '    End Set
    'End Property

    'Private Property _DistinationValue As String
    'Public Property DistinationValue As String
    '    Get
    '        Return _DistinationValue
    '    End Get
    '    Set(value As String)
    '        _DistinationValue = value
    '    End Set
    'End Property


    'Private _IniFile As String
    'Public Property IniFile As String
    '    Get
    '        Return _IniFile
    '    End Get
    '    Set(value As String)
    '        _IniFile = value
    '    End Set
    'End Property

    'Private _ConfigFile As String
    'Public Property ConfigFile As String
    '    Get
    '        Return _ConfigFile
    '    End Get
    '    Set(value As String)
    '        _ConfigFile = value
    '    End Set
    'End Property

    'Private _MyDirectory As String
    'Public Property MyDirectory As String
    '    Get
    '        Return _MyDirectory
    '    End Get
    '    Set(value As String)
    '        _MyDirectory = value
    '    End Set
    'End Property

    'Private Property _ServerSelection As List(Of ServerModel)
    'Public Property ServerSelection As List(Of ServerModel)
    '    Get
    '        Return _ServerSelection
    '    End Get
    '    Set(value As List(Of ServerModel))
    '        _ServerSelection = value
    '    End Set
    'End Property

    'Private Property _DataBaseSelection As List(Of String)
    'Public Property DataBaseSelection As List(Of String)
    '    Get
    '        Return _DataBaseSelection
    '    End Get
    '    Set(value As List(Of String))
    '        _DataBaseSelection = value
    '    End Set
    'End Property

    'Private Property _DataTableSelection As List(Of String)
    'Public Property DataTableSelection As List(Of String)
    '    Get
    '        Return _DataTableSelection
    '    End Get
    '    Set(value As List(Of String))
    '        _DataTableSelection = value
    '    End Set
    'End Property


    'Private _SaveCommand As ICommand
    'Public ReadOnly Property SaveCommand As ICommand
    '    Get
    '        If _SaveCommand Is Nothing Then
    '            _SaveCommand = New DelegateCommand With {
    '                .ExecuteDelegater = AddressOf ,
    '                .CanExecuteDelegater = AddressOf
    '            }
    '        End If
    '    End Get
    'End Property
End Class
