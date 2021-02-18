Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class RpaProjectBase(Of T As {New})
    Inherits Rpa00.JsonHandler(Of T)
    Implements INotifyPropertyChanged

    ' 継承先でプロパティ変更を監視している・・・
    ' 見直した方がいいかも・・・
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Protected Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub

    ' コマンドラインのガイド表示
    Public Overridable ReadOnly Property GuideString As String
        Get
            Return $"{Me.ProjectName}\{Me.RobotAlias}>"
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
    Public MustOverride ReadOnly Property SystemJsonChangedFileName As String
    Public MustOverride ReadOnly Property SystemArchType As Integer
    Public MustOverride ReadOnly Property SystemArchTypeName As String

    Private _MyDirectory As String
    Public Property MyDirectory As String
        Get
            Return Me._MyDirectory
        End Get
        Set(value As String)
            Me._MyDirectory = value
            RaisePropertyChanged("MyDirectory")
        End Set
    End Property

    Private _MyDirectories As List(Of String)
    Public Property MyDirectories As List(Of String)
        Get
            If Me._MyDirectories Is Nothing Then
                Me._MyDirectories = New List(Of String)
            End If
            Return Me._MyDirectories
        End Get
        Set(value As List(Of String))
            Me._MyDirectories = value
        End Set
    End Property

    Private _ProjectName As String
    Public Property ProjectName As String
        Get
            Return Me._ProjectName
        End Get
        Set(value As String)
            Me._ProjectName = value
            RaisePropertyChanged("ProjectName")
        End Set
    End Property

    Private _IsImplicitRobotNameChanged As Boolean
    Private Property IsImplicitRobotNameChanged As Boolean
        Get
            Return Me._IsImplicitRobotNameChanged
        End Get
        Set(value As Boolean)
            Me._IsImplicitRobotNameChanged = value
        End Set
    End Property
    Private _RobotName As String
    Public Property RobotName As String
        Get
            Return Me._RobotName
        End Get
        Set(value As String)
            'Me._RobotAlias = vbNullString
            Me._RobotName = value
            If Not Me.IsImplicitRobotNameChanged Then
                RaisePropertyChanged("RobotName")
                RpaEvent.Instance.RaiseProjectChanged("RobotName")
            End If
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

    <JsonIgnore>
    Public ReadOnly Property MyRobotDirectory As String
        Get
            Return $"{Me.MyDirectory}\{Me.RobotAlias}"
        End Get
    End Property

    <JsonIgnore>
    Public ReadOnly Property MyRobotJsonFileName As String
        Get
            Return $"{Me.MyRobotDirectory}\rpa_project.json"
        End Get
    End Property

    <JsonIgnore>
    Public ReadOnly Property MyRobotReadMeFileName As String
        Get
            Dim [fil] As String = vbNullString
            If Directory.Exists(Me.MyRobotDirectory) Then
                [fil] = $"{Me.MyRobotDirectory}\readme"
            End If
            Return [fil]
        End Get
    End Property

    Private Property _MyRobotReadMe As String
    <JsonIgnore>
    Public Property MyRobotReadMe As String
        Get
            If String.IsNullOrEmpty(Me._MyRobotReadMe) Then
                If File.Exists(Me.MyRobotReadMeFileName) Then
                    Dim i As Integer = GetMyRobotReadMe()
                End If
            End If
            Return Me._MyRobotReadMe
        End Get
        Set(value As String)
            Me._MyRobotReadMe = value
        End Set
    End Property

    ' ロボット本体
    Private Property _MyRobotObject As Object
    <JsonIgnore>
    Public Property MyRobotObject As Object
        Get
            If Me._MyRobotObject Is Nothing Then
                Dim obj = RpaCodes.CreateRobotObject(Me)
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
            RaisePropertyChanged("PrinterName")
        End Set
    End Property

    ' RobotNameアクセサ
    Public Overridable Function SwitchRobot(ByVal [name] As String) As Integer
        Me.RobotName = [name]
        Me.MyRobotObject = Nothing

        Return 0
    End Function

    ' RobotNameアクセサ
    ' このアクセサを使用すると、変更通知を出さない(.chg)
    Public Overridable Function ImplicitSwitchRobot(ByVal [name] As String) As Integer
        Me.IsImplicitRobotNameChanged = True
        Me.RobotName = [name]
        Me.MyRobotObject = Nothing
        Me.IsImplicitRobotNameChanged = False

        Return 0
    End Function

    Public Sub CreateChangedFile(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If Me.FirstLoad Then
            If Not File.Exists(Me.SystemJsonChangedFileName) Then
                RpaModule.CreateChangedFile(Me.SystemJsonChangedFileName)
            End If
        End If
    End Sub

    Public Delegate Function CanExecuteDelegater(ByRef dat As RpaDataWrapper) As Boolean
    Private _CanExecuteHandler As CanExecuteDelegater
    <JsonIgnore>
    Public Property CanExecuteHandler As CanExecuteDelegater
        Get
            Return Me._CanExecuteHandler
        End Get
        Set(value As CanExecuteDelegater)
            Me._CanExecuteHandler = value
        End Set
    End Property

    Public Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        Dim b As Boolean = False
        Try
            Dim dlg As CanExecuteDelegater = Me.CanExecuteHandler
            If dlg Is Nothing Then
                Throw New NotImplementedException
            Else
                b = dlg(dat)
            End If
        Catch ex As Exception
            Console.WriteLine($"({Me.GetType.Name}.CanExecute) {ex.Message}")
            Console.WriteLine()
            b = False
        End Try
        Return b
    End Function

    Private Function GetMyRobotReadMe() As Integer
        Dim txt As String
        Dim sr As StreamReader
        Try
            sr = New System.IO.StreamReader(
                Me.MyRobotReadMeFileName, System.Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING))
            Me.MyRobotReadMe = sr.ReadToEnd()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
        End Try
        Return 0
    End Function

    Public Function DeepCopy() As Object
        Return MemberwiseClone()
    End Function
End Class
