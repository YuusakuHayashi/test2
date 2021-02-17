Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Forms
Imports System.Reflection
Imports System.Collections.ObjectModel

Public Class RpaSystem : Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub

    Private Const RESULT_S = "Success"
    Private Const RESULT_F = "Failed"
    Private Const RESULT_E = "Exception"

    Public Function InitializeData(ByRef dat As Object) As RpaDataWrapper
        Return dat
    End Function

    ' セーブ関連
    '---------------------------------------------------------------------------------------------'
    Private ReadOnly Property CommandTextLogFileName As String
        Get
            Dim [fil] As String = $"{RpaCui.SystemDirectory}\command.log"
            Dim jh As New Rpa00.JsonHandler(Of List(Of RpaCommand))
            If Not File.Exists([fil]) Then
                Call jh.Save(Of List(Of RpaCommand))([fil], (New List(Of RpaCommand)))
            End If
            Return [fil]
        End Get
    End Property

    Private ReadOnly Property SystemLogsDirectory As String
        Get
            Dim [dir] As String = $"{RpaCui.SystemDirectory}\logs"
            If Not Directory.Exists([dir]) Then
                Directory.CreateDirectory([dir])
            End If
            Return [dir]
        End Get
    End Property

    Private _SystemMonthlyLogDirectories As String()
    Private ReadOnly Property SystemMonthlyLogDirectories As String()
        Get
            Dim dirs(12) As String
            If Me._SystemMonthlyLogDirectories Is Nothing Then
                dirs(0) = vbNullString
                For i As Integer = 1 To 12
                    dirs(i) = $"{Me.SystemLogsDirectory}\{String.Format("{0:00}", i)}"
                    If Not Directory.Exists(dirs(i)) Then
                        Directory.CreateDirectory(dirs(i))
                    End If
                Next
                Me._SystemMonthlyLogDirectories = dirs
            End If
            Return Me._SystemMonthlyLogDirectories
        End Get
    End Property

    Private _SystemMonthlyLogFiles As String()
    Private ReadOnly Property SystemMonthlyLogFiles As String()
        Get
            Dim fils(12) As String
            If Me._SystemMonthlyLogFiles Is Nothing Then
                fils(0) = vbNullString
                For i As Integer = 1 To 12
                    fils(i) = $"{Me.SystemMonthlyLogDirectories(i)}\commands.log"
                Next
                Me._SystemMonthlyLogFiles = fils
            End If
            Return Me._SystemMonthlyLogFiles
        End Get
    End Property
    '---------------------------------------------------------------------------------------------'

    Public Enum ExecuteModeNumber
        RpaCui = 0
        RpaGui = 1
    End Enum

    Public Enum RpaReadLineMode
        YesNo = 0
    End Enum

    Private _CommandObject As Object
    Public Property CommandObject As Object
        Get
            Return Me._CommandObject
        End Get
        Set(value As Object)
            Me._CommandObject = value
        End Set
    End Property

    ' 現状、コマンド実行状態しかない
    ' 0 ... 通常
    ' 1 ... コマンド実行中
    Private _StatusCode As Integer
    Public Property StatusCode As Integer
        Get
            Return Me._StatusCode
        End Get
        Set(value As Integer)
            Me._StatusCode = value
            RaisePropertyChanged("StatusCode")
        End Set
    End Property

    ' 2021/02/08 : Gui化のため追加
    Private _OutputText As String
    Public Property OutputText As String
        Get
            Return Me._OutputText
        End Get
        Set(value As String)
            Me._OutputText = value
            RaisePropertyChanged("OutputText")
        End Set
    End Property

    Private _CommandData As RpaCommand
    Public Property CommandData As RpaCommand
        Get
            If Me._CommandData Is Nothing Then
                Me._CommandData = New RpaCommand
            End If
            Return Me._CommandData
        End Get
        Set(value As RpaCommand)
            Me._CommandData = value
        End Set
    End Property

    Private _CommandLogs As ObservableCollection(Of RpaCommand)
    Public Property CommandLogs As ObservableCollection(Of RpaCommand)
        Get
            If Me._CommandLogs Is Nothing Then
                Me._CommandLogs = New ObservableCollection(Of RpaCommand)
            End If
            Return Me._CommandLogs
        End Get
        Set(value As ObservableCollection(Of RpaCommand))
            Me._CommandLogs = value
        End Set
    End Property

    Private _ReturnCode As Integer
    Public Property ReturnCode As Integer
        Get
            Return Me._ReturnCode
        End Get
        Set(value As Integer)
            Me._ReturnCode = value
            RaisePropertyChanged("ReturnCode")
        End Set
    End Property

    Private _ExecuteMode As Integer
    Public Property ExecuteMode As Integer
        Get
            Return Me._ExecuteMode
        End Get
        Set(value As Integer)
            Me._ExecuteMode = value
        End Set
    End Property

    Private _IsRunTime As Boolean
    Public Property IsRunTime As Boolean
        Get
            Return Me._IsRunTime
        End Get
        Set(value As Boolean)
            Me._IsRunTime = value
            RaisePropertyChanged("IsRunTime")
        End Set
    End Property

    Private _ExitFlag As Boolean
    Public Property ExitFlag As Boolean
        Get
            Return Me._ExitFlag
        End Get
        Set(value As Boolean)
            Me._ExitFlag = value
        End Set
    End Property

    ' Ｇｕｉ起動コマンド
    Private _GuiCommandText As String
    Public Property GuiCommandText As String
        Get
            Return Me._GuiCommandText
        End Get
        Set(value As String)
            Me._GuiCommandText = value
        End Set
    End Property

    ' アップデート後コマンドをファストコマンドにするためのロジック(CreateFastBindingCommand())を
    ' 通過したかの判定フラグ
    Private IsUpdatedBindingCommandsCreated As Boolean

    ' １度でもファストコマンドが実行されたかどうかのフラグ
    Private IsFastBindingCommandsExecuted As Boolean

    Private _IsUpdatedBindingCommandsFailed As Boolean
    Private Property IsUpdatedBindingCommandsFailed As Boolean
        Get
            Return Me._IsUpdatedBindingCommandsFailed
        End Get
        Set(value As Boolean)
            Me._IsUpdatedBindingCommandsFailed = value
        End Set
    End Property

    Private _FastBindingCommands As List(Of String)
    Private Property FastBindingCommands As List(Of String)
        Get
            If Me._FastBindingCommands Is Nothing Then
                Me._FastBindingCommands = New List(Of String)
            End If
            Return Me._FastBindingCommands
        End Get
        Set(value As List(Of String))
            Me._FastBindingCommands = value
        End Set
    End Property

    Private _LateBindingCommands As List(Of String)
    Public Property LateBindingCommands As List(Of String)
        Get
            If Me._LateBindingCommands Is Nothing Then
                Me._LateBindingCommands = New List(Of String)
            End If
            Return Me._LateBindingCommands
        End Get
        Set(value As List(Of String))
            Me._LateBindingCommands = value
        End Set
    End Property

    Private _ChainBindingCommands As List(Of String)
    Private Property ChainBindingCommands As List(Of String)
        Get
            If Me._ChainBindingCommands Is Nothing Then
                Me._ChainBindingCommands = New List(Of String)
            End If
            Return Me._ChainBindingCommands
        End Get
        Set(value As List(Of String))
            Me._ChainBindingCommands = value
        End Set
    End Property

    Public ReadOnly Property RpaObject(ByVal index As Integer) As Object
        Get
            Dim ics As New IntranetClientServerProject
            Dim sap As New StandAloneProject
            Dim csp As New ClientServerProject
            Select Case index
                Case ics.SystemArchType : Return ics
                Case sap.SystemArchType : Return sap
                Case csp.SystemArchType : Return csp
                Case Else : Return Nothing
            End Select
        End Get
    End Property

    Public Function LoadCurrentRpa(ByRef dat As Object) As Object
        If Not dat.Initializer.AutoLoad Then
            Return Nothing
        End If
        If dat.Initializer.CurrentProject Is Nothing Then
            Console.WriteLine($"CurrentProject がないため、ロードしませんでした")
            Console.WriteLine()
            Return Nothing
        End If

        Dim rpa = RpaObject(dat.Initializer.CurrentProject.Architecture)
        rpa = rpa.Load(dat.Initializer.CurrentProject.JsonFileName)
        If rpa Is Nothing Then
            Console.WriteLine($"ファイル '{dat.Initializer.CurrentProject.JsonFileName}' のロードに失敗しました")
            Console.WriteLine()
            Return Nothing
        Else
            Console.WriteLine($"ファイル '{dat.Initializer.CurrentProject.JsonFileName}' をロードしました")
            Console.WriteLine()
            Return rpa
        End If
    End Function


    Private _CommandDictionary As Dictionary(Of String, Object)

    Friend ReadOnly Property CommandDictionary As Dictionary(Of String, Object)
        Get
            If Me._CommandDictionary Is Nothing Then
                Me._CommandDictionary = New Dictionary(Of String, Object)
                Me._CommandDictionary.Add("clearscreen", (New ClearScreenCommand))
                Me._CommandDictionary.Add("openexplorer", (New OpenExplorerCommand))
                Me._CommandDictionary.Add("load", (New LoadCommand))
                Me._CommandDictionary.Add("exit", (New ExitCommand))
                Me._CommandDictionary.Add("showmycommands", (New ShowMyCommandCommand))
                Me._CommandDictionary.Add("exportmycommands", (New ExportMyCommandsCommand))
                Me._CommandDictionary.Add("importmycommands", (New ImportMyCommandsCommand))
                Me._CommandDictionary.Add("changecommandenabled", (New ChangeCommandEnabledCommand))
                Me._CommandDictionary.Add("setup", (New SetupRpaCommand))
                Me._CommandDictionary.Add("saveproject", (New SaveProjectCommand))
                Me._CommandDictionary.Add("saveinitializer", (New SaveInitializerCommand))
                Me._CommandDictionary.Add("allsave", (New AllSaveCommand))
                Me._CommandDictionary.Add("runrobot", (New RunRobotCommand))
                Me._CommandDictionary.Add("clonerobot", (New CloneRobotCommand))
                Me._CommandDictionary.Add("attachrobot", (New AttachRobotCommand))
                Me._CommandDictionary.Add("addutility", (New AddUtilityCommand))
                Me._CommandDictionary.Add("addcommandalias", (New AddCommandAliasCommand))
                Me._CommandDictionary.Add("addrobotalias", (New AddRobotAliasCommand))
                Me._CommandDictionary.Add("removerobot", (New RemoveRobotCommand))
                Me._CommandDictionary.Add("removeproject", (New RemoveProjectCommand))
                Me._CommandDictionary.Add("newproject", (New NewProjectCommand))
                Me._CommandDictionary.Add("changeprojectproperty", (New ChangeProjectPropertyCommand))
                Me._CommandDictionary.Add("showprojectproperties", (New ShowProjectPropertiesCommand))
                Me._CommandDictionary.Add("changeprojectpropertyusingfolderbrowser", (New ChangeProjectPropertyUsingFolderBrowserCommand))
                Me._CommandDictionary.Add("showinitializerproperties", (New ShowInitializerPropertiesCommand))
                Me._CommandDictionary.Add("changeinitializerproperty", (New ChangeInitializerPropertyCommand))
                Me._CommandDictionary.Add("updateproject", (New UpdateProjectCommand))
                Me._CommandDictionary.Add("updatemyrobot", (New UpdateMyRobotCommand))
                Me._CommandDictionary.Add("exportrootrobotjson", (New ExportRootRobotJsonCommand))
                Me._CommandDictionary.Add("removemycommand", (New RemoveMyCommandCommand))
                Me._CommandDictionary.Add("activatemycommand", (New ActivateMyCommandCommand))
                Me._CommandDictionary.Add("showmyrobotjson", (New ShowMyRobotJsonCommand))
                Me._CommandDictionary.Add("showmyrobotproperties", (New ShowMyRobotPropertiesCommand))
                Me._CommandDictionary.Add("pushrobot", (New PushRobotCommand))
                Me._CommandDictionary.Add("showrobotreadme", (New ShowRobotReadMeCommand))
                Me._CommandDictionary.Add("help", (New HelpCommand))
                Me._CommandDictionary.Add("selectmydirectory", (New SelectMyDirectoryCommand))
                Me._CommandDictionary.Add("addalias", (New AddAliasCommand))
                Me._CommandDictionary.Add("checkrobot", (New CheckRobotCommand))
            End If
            Return Me._CommandDictionary
        End Get
    End Property

    ' 機能はここに追加
    ' 非常にごちゃごちゃしているので、後々見直し予定・・・
    '---------------------------------------------------------------------------------------------'
    '---------------------------------------------------------------------------------------------'

    Private _CommandBuilders As List(Of Action(Of RpaDataWrapper))
    Private ReadOnly Property CommandBuilders As List(Of Action(Of RpaDataWrapper))
        Get
            If Me._CommandBuilders Is Nothing Then
                Me._CommandBuilders = New List(Of Action(Of RpaDataWrapper))
                Me._CommandBuilders.Add(AddressOf CreateCommandBaseData)
                Me._CommandBuilders.Add(AddressOf CreateCommandProjectData)
                Me._CommandBuilders.Add(AddressOf CreateChainBindingCommands)
                Me._CommandBuilders.Add(AddressOf ConvertAlias)
                Me._CommandBuilders.Add(AddressOf CreateMainCommand)
                Me._CommandBuilders.Add(AddressOf ConvertCommandAlias)
                Me._CommandBuilders.Add(AddressOf CreateUtilityCommand)
                Me._CommandBuilders.Add(AddressOf CreateCommandParameters)
                Me._CommandBuilders.Add(AddressOf CreateCommandObject)
                Me._CommandBuilders.Add(AddressOf CheckCommand)
                Me._CommandBuilders.Add(AddressOf ExecuteCommand)
            End If
            Return Me._CommandBuilders
        End Get
    End Property

    '---------------------------------------------------------------------------------------------'
    ' Initializer等情報の設定
    Private Sub CreateCommandBaseData(dat As RpaDataWrapper)
        Me.CommandData.UserName = dat.Initializer.UserName
        Me.CommandData.RunDate = DateTime.Today
    End Sub

    ' Project情報の設定
    Private Sub CreateCommandProjectData(dat As RpaDataWrapper)
        If dat.Project IsNot Nothing Then
            Me.CommandData.ProjectArchTypeName = dat.Project.SystemArchTypeName
            Me.CommandData.ProjectName = dat.Project.ProjectName
            Me.CommandData.RobotName = dat.Project.RobotName
        End If
    End Sub

    ' BindingCommandの生成
    Private Sub CreateChainBindingCommands(dat As RpaDataWrapper)
        If Me.ChainBindingCommands.Count > 0 Then
            Exit Sub
        End If

        Dim cmds As List(Of String) = Me.CommandData.CommandText.Split("&"c).ToList
        If cmds.Count <= 1 Then
            Exit Sub
        End If

        ' List(Of T).ForEachメソッドが機能してない気がする
        For Each cmd In cmds
            cmd = cmd.Trim
            Me.ChainBindingCommands.Add(cmd)
        Next

        Me.CommandData.CommandText = Me.ChainBindingCommands(0)
        Me.ChainBindingCommands = RpaModule.Pop(Of List(Of String))(Me.ChainBindingCommands)
    End Sub

    ' Aliasの変換
    Private Sub ConvertAlias(dat As RpaDataWrapper)
        Dim [key] As String = Me.CommandData.CommandText
        Dim [keys] As List(Of String) = dat.Initializer.AliasDictionary.Keys.ToList

        If [keys].Contains([key]) Then
            Me.CommandData.CommandText = dat.Initializer.AliasDictionary([key])
        End If
    End Sub

    ' MainCommand の生成
    Private Sub CreateMainCommand(dat As RpaDataWrapper)
        Dim texts() As String

        texts = Me.CommandData.CommandText.Split(" "c)
        Me.CommandData.MainCommand = texts(0)
    End Sub

    ' CommandAlias変換 (TrueCommand の生成)
    Private Sub ConvertCommandAlias(dat As RpaDataWrapper)
        Dim [key] As String = Me.CommandData.CommandText
        Dim [keys] As List(Of String) = dat.Initializer.MyCommandDictionary.Keys.ToList

        ' 初期値
        Me.CommandData.TrueCommand = Me.CommandData.MainCommand

        If [keys].Contains([key]) Then
            If dat.Initializer.MyCommandDictionary([key]).IsEnabled Then
                Me.CommandData.TrueCommand = dat.Initializer.MyCommandDictionary([key]).TrueCommand
            End If
        End If

        ' 自動コマンドの場合、正式名称をセットする
        If Me.CommandData.IsAutoCommand Then
            Me.CommandData.TrueCommand = Me.CommandData.MainCommand
        End If
    End Sub

    ' UtilityCommand の生成
    ' UtilityCommand の場合、 Transaction.TrueCommand も再生成する
    Private Sub CreateUtilityCommand(dat As RpaDataWrapper)
        Dim texts() As String
        texts = Me.CommandData.CommandText.Split(" ")

        If dat.Project IsNot Nothing Then
            If dat.Project.SystemUtilities.Count > 0 Then
                If dat.Project.SystemUtilities.ContainsKey(Me.CommandData.TrueCommand) Then
                    Me.CommandData.UtilityCommand = Me.CommandData.TrueCommand
                    Me.CommandData.TrueCommand = texts(1)
                End If
            End If
        End If
    End Sub

    ' Parameters, .ParametersText の生成
    Private Sub CreateCommandParameters(dat As RpaDataWrapper)
        Dim texts() As String
        texts = Me.CommandData.CommandText.Split(" ")

        Dim ex As Integer
        If String.IsNullOrEmpty(Me.CommandData.UtilityCommand) Then
            ex = 0
        Else
            ex = 1
        End If

        Dim cnt As Integer = 0
        For Each p In texts
            If cnt > ex Then
                Me.CommandData.Parameters.Add(p)
                Me.CommandData.ParametersText &= $"{p} "
            End If
            cnt += 1
        Next

        If Not String.IsNullOrEmpty(Me.CommandData.ParametersText) Then
            Me.CommandData.ParametersText = Me.CommandData.ParametersText.TrimEnd(" ")
        End If
    End Sub

    ' CommandObject生成
    Private Sub CreateCommandObject(dat As RpaDataWrapper)
        If String.IsNullOrEmpty(Me.CommandData.UtilityCommand) Then
            Me.CommandObject = Me.CommandDictionary(Me.CommandData.TrueCommand)
        Else
            Me.CommandObject = dat.Project.SystemUtilities(Me.CommandData.UtilityCommand).UtilityObject.CommandHandler(Me.CommandData.TrueCommand)
        End If

        ' コマンド無効化
        If dat.Initializer.MyCommandDictionary.ContainsKey(Me.CommandData.TrueCommand) Then
            If Not dat.Initializer.MyCommandDictionary(Me.CommandData.TrueCommand).IsEnabled Then
                Me.CommandObject = Nothing
            End If
        End If
    End Sub

    ' コマンドのチェック
    Private Sub CheckCommand(dat As RpaDataWrapper)
        Me.CommandData.Result = RESULT_S
        For Each cmdchk In Me.CommandCheckers(dat)
            If Not cmdchk.ExecuteHandler(dat) Then
                Call RpaWriteLine(cmdchk.MessageHandler(dat))
                Call RpaWriteLine()
                Me.CommandData.Result = RESULT_F
                Me.CommandData.Message = cmdchk.MessageHandler(dat)
                Exit For
            End If
        Next
    End Sub

    ' コマンド実行＆実行時間の設定
    Private Sub ExecuteCommand(dat As RpaDataWrapper)
        If Me.CommandData.Result = RESULT_S Then
            Dim [from] As DateTime = DateTime.Now
            Me.ReturnCode = Me.CommandObject.Execute(dat)
            Dim [to] As DateTime = DateTime.Now
            Me.CommandData.ExecuteTime = [to] - [from]
        Else
            Me.CommandData.ExecuteTime = New TimeSpan(0)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------'

    Private Function BackupMonthlyLog(ByRef dat As RpaDataWrapper) As Integer
        Dim sw As StreamWriter
        Try
            Dim def As Date = New Date
            Dim mon As Integer = 0

            If dat.Initializer.LastActiveDate.CompareTo(def) = 0 Then
                dat.Initializer.LastActiveDate = DateTime.Now
            End If

            mon = dat.Initializer.LastActiveDate.Month

            If mon <> DateTime.Now.Month Then
                File.Copy(
                    Me.CommandTextLogFileName,
                    Me.SystemMonthlyLogFiles(mon),
                    True
                )
                sw = New StreamWriter(
                    Me.CommandTextLogFileName,
                    False,
                    Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING)
                )
            End If
        Catch ex As Exception
            Console.WriteLine($"({Me.GetType.Name}) {ex.Message}")
            Console.WriteLine()
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
        End Try

        Return 0
    End Function

    Private Function SendMonthlyLog(ByRef dat As RpaDataWrapper) As Integer
        Try
            If dat.Project Is Nothing Then
                Return 0
            End If

            If dat.Project.SystemArchTypeName <> (New IntranetClientServerProject).GetType.Name Then
                Return 0
            End If

            Dim logdirs() As String = Me.SystemMonthlyLogDirectories
            For i As Integer = LBound(logdirs) To UBound(logdirs)
                Dim logs() As String
                If Directory.Exists(logdirs(i)) Then
                    logs = Directory.GetFiles(logdirs(i))
                Else
                    Continue For
                End If
                For j As Integer = LBound(logs) To UBound(logs)
                    File.Copy(logs(j), $"{dat.Project.RootMonthlyLogDirectories(i)}\{(DateTime.Now).ToString("yyyyMMddHHmmss")}_{j.ToString}.log")
                    File.Delete(logs(j))
                Next
            Next
        Catch ex As Exception
            Console.WriteLine($"({Me.GetType.Name}) {ex.Message}")
            Console.WriteLine()
            Console.ReadLine()
        End Try

        Return 0
    End Function

    Public Class CommandChecker
        Private _MessageHandler As Func(Of RpaDataWrapper, String)
        Public Property MessageHandler As Func(Of RpaDataWrapper, String)
            Get
                Return Me._MessageHandler
            End Get
            Set(value As Func(Of RpaDataWrapper, String))
                Me._MessageHandler = value
            End Set
        End Property

        Private _ExecuteHandler As Func(Of RpaDataWrapper, Boolean)
        Public Property ExecuteHandler As Func(Of RpaDataWrapper, Boolean)
            Get
                Return Me._ExecuteHandler
            End Get
            Set(value As Func(Of RpaDataWrapper, Boolean))
                Me._ExecuteHandler = value
            End Set
        End Property
    End Class

    Private _CommandCheckers As List(Of CommandChecker)
    Private ReadOnly Property CommandCheckers(dat As RpaDataWrapper) As List(Of CommandChecker)
        Get
            If Me._CommandCheckers Is Nothing Then
                Dim cmdchk As CommandChecker

                Me._CommandCheckers = New List(Of CommandChecker)
                Me._CommandCheckers.Add(
                    New CommandChecker With {.ExecuteHandler = AddressOf CommandCheckExecuter1, .MessageHandler = AddressOf CommandCheckMessenger1}
                )
                Me._CommandCheckers.Add(
                    New CommandChecker With {.ExecuteHandler = AddressOf CommandCheckExecuter2, .MessageHandler = AddressOf CommandCheckMessenger2}
                )
                Me._CommandCheckers.Add(
                    New CommandChecker With {.ExecuteHandler = AddressOf CommandCheckExecuter3, .MessageHandler = AddressOf CommandCheckMessenger3}
                )
                Me._CommandCheckers.Add(
                    New CommandChecker With {.ExecuteHandler = AddressOf CommandCheckExecuter4, .MessageHandler = AddressOf CommandCheckMessenger4}
                )
                Me._CommandCheckers.Add(
                    New CommandChecker With {.ExecuteHandler = AddressOf CommandCheckExecuter5, .MessageHandler = AddressOf CommandCheckMessenger5}
                )
                Me._CommandCheckers.Add(
                    New CommandChecker With {.ExecuteHandler = AddressOf CommandCheckExecuter6, .MessageHandler = AddressOf CommandCheckMessenger6}
                )
            End If
            Return Me._CommandCheckers
        End Get
    End Property

    Private Function CommandCheckExecuter1(dat As RpaDataWrapper) As Boolean
        If Me.CommandObject Is Nothing Then
            Return False
        End If
        Return True
    End Function
    Private Function CommandCheckMessenger1(dat As RpaDataWrapper) As String
        Return $"コマンド  '{Me.CommandData.CommandText}' はありません"
    End Function

    Private Function CommandCheckExecuter2(dat As RpaDataWrapper) As Boolean
        If Me.CommandObject.ExecutableParameterCount(0) > Me.CommandData.Parameters.Count _
        Or Me.CommandObject.ExecutableParameterCount(1) < Me.CommandData.Parameters.Count Then
            Return False
        End If
        Return True
    End Function
    Private Function CommandCheckMessenger2(dat As RpaDataWrapper) As String
        Return $"パラメータ数が範囲外です: 許容パラメータ数範囲={Me.CommandObject.ExecutableParameterCount(0)}~{Me.CommandObject.ExecutableParameterCount(1)}"
    End Function

    Private Function CommandCheckExecuter3(dat As RpaDataWrapper) As Boolean
        If Not Me.CommandObject.ExecuteIfNoProject Then
            If dat.Project Is Nothing Then
                Return False
            End If
        End If
        Return True
    End Function
    Private Function CommandCheckMessenger3(dat As RpaDataWrapper) As String
        Return $"プロジェクトが存在しないため、実行できません"
    End Function

    Private Function CommandCheckExecuter4(dat As RpaDataWrapper) As Boolean
        Dim archs As String() = Me.CommandObject.ExecutableProjectArchitectures
        If Not archs.Contains("AllArchitectures") Then
            If Not archs.Contains(dat.Project.SystemArchTypeName) Then
                Return False
            End If
        End If
        Return True
    End Function
    Private Function CommandCheckMessenger4(dat As RpaDataWrapper) As String
        Return $"このプロジェクト構成では、コマンド '{Me.CommandData.CommandText}' は実行できません"
    End Function

    Private Function CommandCheckExecuter5(dat As RpaDataWrapper) As Boolean
        Dim users As String() = Me.CommandObject.ExecutableUser
        If Not users.Contains("AllUser") Then
            If Not users.Contains(dat.Initializer.UserLevel) Then
                Return False
            End If
        End If
        Return True
    End Function
    Private Function CommandCheckMessenger5(dat As RpaDataWrapper) As String
        Return $"このユーザでは、コマンド '{Me.CommandData.CommandText}' は実行できません"
    End Function

    Private Function CommandCheckExecuter6(dat As RpaDataWrapper) As Boolean
        If Not Me.CommandObject.CanExecute(dat) Then
            Return False
        End If
        Return True
    End Function
    Private Function CommandCheckMessenger6(dat As RpaDataWrapper) As String
        Return $"コマンドは実行出来ませんでした"
    End Function

    Private Sub CommandLoop(ByRef dat As RpaDataWrapper)
        Do
            Me.CommandData = New RpaCommand
            'Dim cmdlog As RpaTransaction.CommandLogData = Nothing
            Dim fastcmdflag As Boolean = False
            Dim latecmdflag As Boolean = False
            Try
                Me.IsRunTime = True
                Dim autocmdflag As Boolean = False

                If Me.ExecuteMode = ExecuteModeNumber.RpaCui Then
                    ' 一度のみ実行
                    If Not Me.IsUpdatedBindingCommandsCreated Then
                        Dim a As Integer = CreateFastBindingCommands(dat)
                    End If

                    If Me.LateBindingCommands.Count > 0 Then
                        Me.CommandData.IsAutoCommand = True
                        autocmdflag = True
                        latecmdflag = True
                        Me.CommandData.CommandText = Me.LateBindingCommands(0)
                        Me.LateBindingCommands = RpaModule.Pop(Of List(Of String))(Me.LateBindingCommands)
                    ElseIf Me.FastBindingCommands.Count > 0 Then
                        Call RpaWriteLine($"更新プログラムを実行中・・・")
                        Me.CommandData.IsAutoCommand = True
                        autocmdflag = True
                        fastcmdflag = True
                        If Not IsFastBindingCommandsExecuted Then
                            IsFastBindingCommandsExecuted = True
                        End If
                        Me.CommandData.CommandText = Me.FastBindingCommands(0)
                        Me.FastBindingCommands = RpaModule.Pop(Of List(Of String))(Me.FastBindingCommands)
                    ElseIf Me.ChainBindingCommands.Count > 0 Then
                        Me.CommandData.CommandText = Me.ChainBindingCommands(0)
                        Me.ChainBindingCommands = RpaModule.Pop(Of List(Of String))(Me.ChainBindingCommands)
                    Else
                        Me.CommandData.CommandText = dat.Transaction.ShowRpaIndicator(dat)
                    End If
                End If
                If Me.ExecuteMode = ExecuteModeNumber.RpaGui Then
                    Me.CommandData.CommandText = Me.GuiCommandText
                End If

                Me.StatusCode = 1

                ' コマンド生成
                '---------------------------------------------------------------------------------'
                For Each builder In Me.CommandBuilders
                    Call builder(dat)
                Next
                '---------------------------------------------------------------------------------'

            Catch ex As Exception
                ' Exceptionを利用するのは気持ち悪いので、修正するかも・・・
                ' Me.CommandData.Result が空文字の場合、例外が発生する場合
                Call RpaWriteLine($"({Me.GetType.Name}) {ex.Message}")
                Call RpaWriteLine()
                Me.CommandData.Result = RESULT_E
                Me.CommandData.Message = ex.Message

                If latecmdflag Or fastcmdflag Then
                    Call RpaWriteLine($"自動生成されたコマンド '{Me.CommandData.CommandText}' は失敗しました")
                    Call RpaWriteLine()
                    Me.LateBindingCommands = New List(Of String)
                    Me.FastBindingCommands = New List(Of String)
                End If

                If fastcmdflag Then
                    ' 今のところファストコマンドはアップデート後コマンドしかないので通用するが・・・
                    Me.IsUpdatedBindingCommandsFailed = True
                End If
            Finally
                ' 今のところファストコマンドはアップデート後コマンドしかないので通用するが・・・
                If Me.IsFastBindingCommandsExecuted And Me.FastBindingCommands.Count = 0 And (Not Me.IsUpdatedBindingCommandsFailed) Then
                    Dim a As Integer = CompleteUpdate(dat)
                End If

                '-------------------------------------------------------------------'
                Me.CommandLogs.Add(Me.CommandData)
                'MyEventHandler.Instance.RaiseCommandLogsChanged(Me.CommandLogs)
                '-------------------------------------------------------------------'

                'If Me.ExecuteMode = ExecuteModeNumber.RpaGui Then
                '    Me.ExitFlag = True
                'End If
                ' ExitFlag のチェック
                ' ExitFlag が立った場合でも、LateBindingCommandを処理する必要がある
                'Call CheckExitFlag(dat)

                Me.StatusCode = 0
                Me.IsRunTime = False
            End Try

            ' 非対応のコマンドが多すぎるため
            If Me.ExecuteMode = ExecuteModeNumber.RpaGui Then
                Exit Do
            End If

            If Me.ExitFlag And Me.FastBindingCommands.Count = 0 And Me.LateBindingCommands.Count = 0 And Me.ChainBindingCommands.Count = 0 Then
                Exit Do
            End If
        Loop Until False
    End Sub

    'Private Sub CheckExitFlag(ByRef dat As RpaDataWrapper)
    '    If Me.ExitFlag Then
    '        Me.LateExitFlag = True
    '    End If
    '    If Me.LateExitFlag And Me.FastBindingCommands.Count = 0 And Me.LateBindingCommands.Count = 0 And Me.ChainBindingCommands.Count = 0 Then
    '        Me.ExitFlag = True
    '    Else
    '        Me.ExitFlag = False
    '    End If
    'End Sub

    Public Sub Main(ByRef dat As RpaDataWrapper)
        Dim i As Integer = BackupMonthlyLog(dat)               ' 当月ログ保存

        Call CommandLoop(dat)


        ' プロジェクトセーブ
        '-----------------------------------------------------------------------------------------'
        If dat.Project IsNot Nothing Then
            If File.Exists(dat.Project.SystemJsonChangedFileName) Then
                Dim yorn As String = vbNullString
                Do
                    yorn = vbNullString
                    Console.WriteLine()
                    yorn = RpaReadLine(dat, RpaReadLineMode.YesNo, $"Projectに変更があります。変更を保存しますか？ (y/n)")
                    If yorn = "y" Or yorn = "n" Then
                        Exit Do
                    End If
                Loop Until False
                If yorn = "y" Then
                    Call Me.Save(dat.Project.SystemJsonFileName, dat.Project, dat.Project.SystemJsonChangedFileName)
                Else
                    File.Delete(dat.Project.SystemJsonChangedFileName)
                End If
            End If
        End If
        '-----------------------------------------------------------------------------------------'

        ' イニシャライザセーブ
        '-----------------------------------------------------------------------------------------'
        If File.Exists(RpaInitializer.SystemIniChangedFileName) Then
            Dim yorn As String = vbNullString
            Do
                yorn = vbNullString
                Console.WriteLine()
                yorn = RpaReadLine(dat, RpaReadLineMode.YesNo, $"Initializerに変更があります。変更を保存しますか？ (y/n)")
                If yorn = "y" Or yorn = "n" Then
                    Exit Do
                End If
            Loop Until False
            If yorn = "y" Then
                Call Me.Save(RpaCui.SystemIniFileName, dat.Initializer, RpaInitializer.SystemIniChangedFileName)
            Else
                File.Delete(RpaInitializer.SystemIniChangedFileName)
            End If
        End If
        '-----------------------------------------------------------------------------------------'

        ' ログセーブ
        Call SaveCommandLogs()

        ' Cuiはコンソールを閉じるとき、このロジックを通過するため、問題がないが
        ' Guiは何度もMain()を呼ぶ可能性がある。このため、ログは都度削除する
        Me.CommandLogs.RemoveAt(0)

        Dim k As Integer = SendMonthlyLog(dat)    ' 当月ログ送信
    End Sub

    Public Sub SaveCommandLogs()
        Dim sw As New StreamWriter(
            Me.CommandTextLogFileName,
            True,
            Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING)
        )

        For Each [log] In Me.CommandLogs
            Dim props As PropertyInfo() = [log].GetType.GetProperties()
            Dim txt As String = vbNullString
            For Each prop In props
                If prop.CanWrite Then
                    If prop.GetValue([log]) IsNot Nothing Then
                        txt &= $"{prop.GetValue([log]).ToString()},"
                    Else
                        txt &= $"{vbNullString},"
                    End If
                End If
            Next
            sw.WriteLine(txt)
        Next

        sw.Close()
        sw.Dispose()
    End Sub


    Private Function CreateFastBindingCommands(ByRef dat As RpaDataWrapper) As Integer
        If dat.Project Is Nothing Then
            Return 0
        End If
        If dat.Project.SystemArchTypeName <> (New IntranetClientServerProject).GetType.Name Then
            Return 0
        End If

        Me.IsUpdatedBindingCommandsCreated = True

        Dim jh As New Rpa00.JsonHandler(Of List(Of RpaUpdater))
        Dim srus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile)
        Dim srus2 As List(Of RpaUpdater)
        srus2 = srus.FindAll(
            Function(sru)
                Return (sru.UpdatedBindingCommands.Count > 0)
            End Function
        )
        srus2.Sort(
            Function(before, after)
                Return (before.ReleaseDate < after.ReleaseDate)
            End Function
        )

        If srus2.Count = 0 Then
            Return 0
        End If

        If srus2.Last.UpdaterProcessId = dat.Initializer.ProcessId Then
            Return 0
        End If

        ' コマンド生成
        For Each sru In srus2
            For Each ubc In sru.UpdatedBindingCommands
                Me.FastBindingCommands.Add(ubc)
            Next
        Next

        Return 0
    End Function

    Private Function CompleteUpdate(ByRef dat As RpaDataWrapper) As Integer
        Dim jh As New Rpa00.JsonHandler(Of List(Of RpaUpdater))
        Dim srus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile)
        srus.Sort(
            Function(before, after)
                Return (before.ReleaseDate < after.ReleaseDate)
            End Function
        )
        For Each sru In srus
            If sru.UpdatedBindingCommands.Count > 0 Then
                sru.UpdatedBindingCommands = New List(Of String)
            End If
        Next
        Call jh.Save(Of List(Of RpaUpdater))(dat.Project.SystemRobotsUpdateFile, srus)

        Return 0
    End Function

    Public Sub RpaWriteLine(Optional wline As String = vbNullString)
        ' Cuiの場合
        If Me.ExecuteMode = ExecuteModeNumber.RpaCui Then
            Console.WriteLine(wline)
        ElseIf Me.ExecuteMode = ExecuteModeNumber.RpaGui Then
            Me.OutputText &= $"{wline}{vbCrLf}"
        End If
    End Sub

    Public Function RpaReadLine(ByRef dat As RpaDataWrapper, ByVal mode As Integer, Optional wline As String = vbNullString) As String
        ' Cuiの場合
        Call RpaWriteLine(wline)
        If Me.ExecuteMode = ExecuteModeNumber.RpaCui Then
            If dat.Project IsNot Nothing Then
                Console.Write($"{dat.Project.GuideString}")
            Else
                Console.Write("NoRpa>")
            End If
            Return Console.ReadLine()
        End If
        If Me.ExecuteMode = ExecuteModeNumber.RpaGui Then
            If mode = RpaReadLineMode.YesNo Then
                Dim dr As DialogResult = MessageBox.Show(wline, vbNullString, MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2)
                If dr = DialogResult.Yes Then
                    Return "y"
                End If
                If dr = DialogResult.No Then
                    Return "n"
                End If
            End If
        End If
    End Function

    Public Sub Save(ByVal savefile As String, ByRef obj As Object, ByVal chgfile As String)
        obj.Save(savefile, obj)
        File.Delete(chgfile)
        Call Me.RpaWriteLine($"ファイル '{savefile} をセーブしました")
        Call Me.RpaWriteLine()
    End Sub

    'Private Async Sub AsyncRpaWriteLine(ByVal wline As String)
    '    ' Guiの場合
    '    Dim t As Task = Task.Run(
    '        Sub()
    '            Me.Output.OutputText &= $"{wline}{vbCrLf}"
    '        End Sub
    '    )
    '    Await t
    'End Sub
End Class
