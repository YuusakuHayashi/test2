Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class RpaProjectBase(Of T As {New})
    Inherits JsonHandler(Of T)

    Private _UserName As String
    Public Property UserName As String
        Get
            Return Me._UserName
        End Get
        Set(value As String)
            Me._UserName = value
        End Set
    End Property

    <JsonIgnore>
    Public Shared ReadOnly Property SystemDirectory As String
        Get
            Return System.Environment.GetEnvironmentVariable("USERPROFILE") & "\rpa_project"
        End Get
    End Property

    <JsonIgnore>
    Public Shared ReadOnly Property SystemUtilityDirectory As String
        Get
            Dim d = $"{RpaProjectBase(Of T).SystemDirectory}\util"
            If Not Directory.Exists(d) Then
                Directory.CreateDirectory(d)
            End If
            Return d
        End Get
    End Property

    <JsonIgnore>
    Public Shared ReadOnly Property SystemIniFileName As String
        Get
            Return $"{RpaProjectBase(Of T).SystemDirectory}\rpa.ini"
        End Get
    End Property

    Private _SystemUtilities As Dictionary(Of String, RpaUtility)
    Public Property SystemUtilities As Dictionary(Of String, RpaUtility)
        Get
            If Me._SystemUtilities Is Nothing Then
                Me._SystemUtilities = New Dictionary(Of String, RpaUtility)
            End If
            Return Me._SystemUtilities
        End Get
        Set(value As Dictionary(Of String, RpaUtility))
            Me._SystemUtilities = value
        End Set
    End Property

    Public MustOverride ReadOnly Property SystemArchitecutureDirectory As String

    Public MustOverride ReadOnly Property SystemSolutionDirectory As String

    Public MustOverride ReadOnly Property SystemJsonFileName As String

    <JsonIgnore>
    Public Shared ReadOnly Property SystemDllDirectory As String
        Get
            Return $"{RpaProjectBase(Of T).SystemDirectory}\dll"
        End Get
    End Property

    <JsonIgnore>
    Public Shared ReadOnly Property System00DllFileName As String
        Get
            Return $"{RpaProjectBase(Of T).SystemDllDirectory}\Rpa00.dll"
        End Get
    End Property

    Private _MyDirectory As String
    Public Property MyDirectory As String
        Get
            Return Me._MyDirectory
        End Get
        Set(value As String)
            Me._MyDirectory = value
        End Set
    End Property

    Private _SolutionName As String
    Public Property SolutionName As String
        Get
            Return Me._SolutionName
        End Get
        Set(value As String)
            Me._SolutionName = value
        End Set
    End Property

    Private _ProjectName As String
    Public Property ProjectName As String
        Get
            Return Me._ProjectName
        End Get
        Set(value As String)
            Me._ProjectAlias = vbNullString
            Me._ProjectName = value
        End Set
    End Property

    Private _AliasDictionary As Dictionary(Of String, String)
    Public Property AliasDictionary As Dictionary(Of String, String)
        Get
            If Me._AliasDictionary Is Nothing Then
                Me._AliasDictionary = New Dictionary(Of String, String)
            End If
            Return Me._AliasDictionary
        End Get
        Set(value As Dictionary(Of String, String))
            Me._AliasDictionary = value
        End Set
    End Property

    Private _ProjectAlias As String
    Public ReadOnly Property ProjectAlias As String
        Get
            If String.IsNullOrEmpty(Me._ProjectAlias) Then
                If String.IsNullOrEmpty(Me.ProjectName) Then
                    Me._ProjectAlias = "No Project"
                Else
                    If Me.AliasDictionary.ContainsKey(Me.ProjectName) Then
                        Me._ProjectAlias = Me.AliasDictionary(Me.ProjectName)
                    Else
                        Me._ProjectAlias = Me.ProjectName
                    End If
                End If
            End If
            Return Me._ProjectAlias
        End Get
    End Property

    Public ReadOnly Property IsMyProjectExists As Boolean
        Get
            If Directory.Exists(Me.MyProjectDirectory) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property MyProjectDirectory As String
        Get
            Return $"{Me.MyDirectory}\{Me.ProjectAlias}"
        End Get
    End Property

    Public ReadOnly Property MyProjectJsonFileName As String
        Get
            Return $"{Me.MyProjectDirectory}\rpa_project.json"
        End Get
    End Property

    Private Property _MyProjectObject As Object
    <JsonIgnore>
    Public Property MyProjectObject As Object
        Get
            If Me._MyProjectObject Is Nothing Then
                Dim obj = RpaCodes.RpaObject(Me)
                Dim obj2 = Nothing
                If obj Is Nothing Then
                    Me._MyProjectObject = Nothing
                Else
                    If File.Exists(Me.MyProjectJsonFileName) Then
                        obj2 = obj.Load(Me.MyProjectJsonFileName)
                    End If
                    If obj2 Is Nothing Then
                        Me._MyProjectObject = obj
                    Else
                        Me._MyProjectObject = obj2
                    End If
                End If
            End If
            Return Me._MyProjectObject
        End Get
        Set(value As Object)
            Me._MyProjectObject = value
        End Set
    End Property

    Private _PrinterName As String
    Public Property PrinterName As String
        Get
            Return Me._PrinterName
        End Get
        Set(value As String)
            Me._PrinterName = value
        End Set
    End Property

    Private _UsePrinterName As String
    <JsonIgnore>
    Public Property UsePrinterName As String
        Get
            Return Me._UsePrinterName
        End Get
        Set(value As String)
            Me._UsePrinterName = value
        End Set
    End Property


    Public MustOverride Sub CheckProject()
End Class
