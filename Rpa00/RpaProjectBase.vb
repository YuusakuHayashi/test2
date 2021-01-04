Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class RpaProjectBase(Of T As {New})
    Inherits JsonHandler(Of T)
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub

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

    Public Class RpaUtility
        Private _UtilityName As String
        Public Property UtilityName As String
            Get
                Return _UtilityName
            End Get
            Set(value As String)
                Me._UtilityName = value
            End Set
        End Property

        Private _UtilityObject As Object
        <JsonIgnore>
        Public Property UtilityObject As Object
            Get
                If Me._UtilityObject Is Nothing Then
                    Me._UtilityObject = RpaCodes.RpaUtilityObject(UtilityName)
                End If
                Return Me._UtilityObject
            End Get
            Set(value As Object)
                Me._UtilityObject = value
            End Set
        End Property
    End Class


    Public MustOverride ReadOnly Property SystemArchDirectory As String
    Public MustOverride ReadOnly Property SystemProjectDirectory As String
    Public MustOverride ReadOnly Property SystemJsonFileName As String
    Public MustOverride ReadOnly Property SystemJsonChangeFileName As String
    Public MustOverride ReadOnly Property SystemArchType As Integer
    Public MustOverride ReadOnly Property SystemArchTypeName As String


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

    Private _ProjectName As String
    Public Property ProjectName As String
        Get
            Return Me._ProjectName
        End Get
        Set(value As String)
            Me._ProjectName = value
        End Set
    End Property

    Private _RobotName As String
    Public Property RobotName As String
        Get
            Return Me._RobotName
        End Get
        Set(value As String)
            Me._RobotAlias = vbNullString
            Me._RobotName = value
        End Set
    End Property

    Private _RobotAliasDictionary As Dictionary(Of String, String)
    Public Property RobotAliasDictionary As Dictionary(Of String, String)
        Get
            If Me._RobotAliasDictionary Is Nothing Then
                Me._RobotAliasDictionary = New Dictionary(Of String, String)
            End If
            Return Me._RobotAliasDictionary
        End Get
        Set(value As Dictionary(Of String, String))
            Me._RobotAliasDictionary = value
        End Set
    End Property

    Private _RobotAlias As String
    Public ReadOnly Property RobotAlias As String
        Get
            If String.IsNullOrEmpty(Me.RobotName) Then
                Me._RobotAlias = "No Robot"
            Else
                If Me.RobotAliasDictionary.ContainsKey(Me.RobotName) Then
                    Me._RobotAlias = Me.RobotAliasDictionary(Me.RobotName)
                Else
                    Me._RobotAlias = Me.RobotName
                End If
            End If
            Return Me._RobotAlias
        End Get
    End Property

    Public ReadOnly Property IsMyRobotExists As Boolean
        Get
            If Directory.Exists(Me.MyRobotDirectory) Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property MyRobotDirectory As String
        Get
            Return $"{Me.MyDirectory}\{Me.RobotAlias}"
        End Get
    End Property

    Public ReadOnly Property MyRobotJsonFileName As String
        Get
            Return $"{Me.MyRobotDirectory}\rpa_project.json"
        End Get
    End Property

    Private Property _MyRobotObject As Object
    <JsonIgnore>
    Public Property MyRobotObject As Object
        Get
            If Me._MyRobotObject Is Nothing Then
                Dim obj = RpaCodes.RpaObject(Me)
                Dim obj2 = Nothing
                If obj Is Nothing Then
                    Me._MyRobotObject = Nothing
                Else
                    If File.Exists(Me.MyRobotJsonFileName) Then
                        obj2 = obj.Load(Me.MyRobotJsonFileName)
                    End If
                    If obj2 Is Nothing Then
                        Me._MyRobotObject = obj
                    Else
                        Me._MyRobotObject = obj2
                    End If
                End If
            End If
            Return Me._MyRobotObject
        End Get
        Set(value As Object)
            Me._MyRobotObject = value
        End Set
    End Property


    Public Sub AddUtility(ByVal util As String)
        Me.SystemUtilities.Add(
            util, (New RpaUtility With {
                .UtilityName = util,
                .UtilityObject = RpaCodes.RpaUtilityObject(util)
            })
        )
    End Sub

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

    Public Sub CreateChangedFile()
        If Me.FirstLoad Then
            RpaModule.CreateChangedFile(Me.SystemJsonChangeFileName)
        End If
    End Sub

    Public Function DeepCopy() As Object
        Return MemberwiseClone()
    End Function
End Class
