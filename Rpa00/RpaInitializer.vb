Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.ComponentModel

Public Class RpaInitializer
    Inherits RpaCui.JsonHandler(Of RpaInitializer)
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub

    '<JsonIgnore>
    'Public Shared ReadOnly Property SystemIniFileName As String
    '    Get
    '        Return $"{RpaCui.SystemDirectory}\rpa.ini"
    '    End Get
    'End Property

    Private _ProcessId As Integer
    <JsonIgnore>
    Public Property ProcessId As Integer
        Get
            Do Until Me._ProcessId > 0
                Me._ProcessId = (New Random).Next
            Loop
            Return Me._ProcessId
        End Get
        Set(value As Integer)
            Me._ProcessId = value
        End Set
    End Property

    Private _LastActiveDate As Date
    Public Property LastActiveDate As Date
        Get
            Return Me._LastActiveDate
        End Get
        Set(value As Date)
            Me._LastActiveDate = value
        End Set
    End Property

    Private _UserName As String
    Public Property UserName As String
        Get
            Return Me._UserName
        End Get
        Set(value As String)
            Me._UserName = value
            RaisePropertyChanged("UserName")
        End Set
    End Property

    Private _UserLevel As String
    Public Property UserLevel As String
        Get
            Return Me._UserLevel
        End Get
        Set(value As String)
            Me._UserLevel = value
            RaisePropertyChanged("UserLevel")
        End Set
    End Property

    Private _ReleaseVersion As String
    Public Property ReleaseVersion As String
        Get
            Return Me._ReleaseVersion
        End Get
        Set(value As String)
            Me._ReleaseVersion = value
            RaisePropertyChanged("ReleaseVersion")
        End Set
    End Property

    <JsonIgnore>
    Public Shared ReadOnly Property SystemIniChangedFileName As String
        Get
            Return $"{RpaCui.SystemDirectory}\rpa.ini.chg"
        End Get
    End Property


    Private _CurrentProject As RpaProject
    Public Property CurrentProject As RpaProject
        Get
            Return Me._CurrentProject
        End Get
        Set(value As RpaProject)
            Me._CurrentProject = value
            RaisePropertyChanged("CurrentProject")
        End Set
    End Property

    Private _AutoLoad As Boolean
    Public Property AutoLoad As Boolean
        Get
            Return Me._AutoLoad
        End Get
        Set(value As Boolean)
            Me._AutoLoad = value
            RaisePropertyChanged("AutoLoad")
        End Set
    End Property

    Private _Projects As List(Of RpaProject)
    Public Property Projects As List(Of RpaProject)
        Get
            If Me._Projects Is Nothing Then
                Me._Projects = New List(Of RpaProject)
            End If
            Return Me._Projects
        End Get
        Set(value As List(Of RpaProject))
            Me._Projects = value
        End Set
    End Property

    Public Class RpaProject
        Private _Architecture As Integer
        Public Property Architecture As Integer
            Get
                Return Me._Architecture
            End Get
            Set(value As Integer)
                Me._Architecture = value
            End Set
        End Property

        Private _Name As String
        Public Property Name As String
            Get
                Return Me._Name
            End Get
            Set(value As String)
                Me._Name = value
            End Set
        End Property

        Private _ProjectDirectory As String
        Public Property ProjectDirectory As String
            Get
                Return Me._ProjectDirectory
            End Get
            Set(value As String)
                Me._ProjectDirectory = value
            End Set
        End Property

        Private _JsonFileName As String
        Public Property JsonFileName As String
            Get
                Return Me._JsonFileName
            End Get
            Set(value As String)
                Me._JsonFileName = value
            End Set
        End Property
    End Class

    Public Shared ReadOnly Property MyCommandsJsonFileName As String
        Get
            Return $"{RpaCui.SystemDirectory}\mycommands.json"
        End Get
    End Property


    Private _MyCommandDictionary As Dictionary(Of String, MyCommand)


    Public Property MyCommandDictionary As Dictionary(Of String, MyCommand)
        Get
            If Me._MyCommandDictionary Is Nothing Then
                Me._MyCommandDictionary = New Dictionary(Of String, MyCommand)
            End If
            Return Me._MyCommandDictionary
        End Get
        Set(value As Dictionary(Of String, MyCommand))
            Me._MyCommandDictionary = value
        End Set
    End Property

    Public Class MyCommand
        Public [Alias] As String
        Public IsEnabled As Boolean
    End Class

    'Public Function BeginTransaction()
    '    Call Save(RpaInitializer.SystemIniTempFileName, Me)
    '    Return Me
    'End Function

    'Public Function TransactionRollBack() As RpaInitializer
    '    Dim ini = Load(RpaInitializer.SystemIniTempFileName)
    '    File.Delete(RpaInitializer.SystemIniTempFileName)
    '    Return ini
    'End Function

    Private Sub CreateChangedFile(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If Me.FirstLoad Then
            If Not File.Exists(RpaInitializer.SystemIniChangedFileName) Then
                Call RpaModule.CreateChangedFile(RpaInitializer.SystemIniChangedFileName)
            End If
        End If
    End Sub

    Sub New()
        AddHandler Me.PropertyChanged, AddressOf CreateChangedFile
    End Sub

    Protected Overrides Sub Finalize()
        RemoveHandler Me.PropertyChanged, AddressOf CreateChangedFile
        MyBase.Finalize()
    End Sub
End Class
