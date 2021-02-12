Imports System.ComponentModel
Imports System.IO

Public Class RpaSystem : Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
    End Sub

    Public Enum ExecuteModeNumber
        RpaCui = 0
        RpaGui = 1
    End Enum

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

    'Private _Output As OutputData
    'Public Property Output As OutputData
    '    Get
    '        If Me._Output Is Nothing Then
    '            Me._Output = New OutputData
    '        End If
    '        Return Me._Output
    '    End Get
    '    Set(value As OutputData)
    '        Me._Output = value
    '    End Set
    'End Property

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

    ' 自動
    Private _LateExitFlag As Boolean
    Private Property LateExitFlag As Boolean
        Get
            Return Me._LateExitFlag
        End Get
        Set(value As Boolean)
            Me._LateExitFlag = value
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
    Private Function CreateCommand(dat As RpaDataWrapper) As Object
        ' コマンド生成
        Dim cmd = Nothing
        If String.IsNullOrEmpty(dat.Transaction.UtilityCommand) Then
            cmd = Me.CommandDictionary(dat.Transaction.TrueCommand)
        Else
            cmd = dat.Project.SystemUtilities(dat.Transaction.UtilityCommand).UtilityObject.CommandHandler(dat.Transaction.TrueCommand)
        End If

        ' コマンド無効化
        If dat.Initializer.MyCommandDictionary.ContainsKey(dat.Transaction.TrueCommand) Then
            If Not dat.Initializer.MyCommandDictionary(dat.Transaction.TrueCommand).IsEnabled Then
                cmd = Nothing
            End If
        End If

        Return cmd
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Function CheckAlias(dat As RpaDataWrapper) As Integer
        Dim [key] As String = dat.Transaction.CommandText
        Dim [keys] As List(Of String) = dat.Initializer.AliasDictionary.Keys.ToList

        If [keys].Contains([key]) Then
            dat.Transaction.CommandText = dat.Initializer.AliasDictionary([key])
        End If

        Return 0
    End Function

    ' 1. Transaction.MainCommand の更新
    Private Function CheckMainCommand(ByRef dat As RpaDataWrapper) As Integer
        Dim texts() As String
        dat.Transaction.MainCommand = vbNullString

        texts = dat.Transaction.CommandText.Split(" ")
        dat.Transaction.MainCommand = texts(0)

        Return 0
    End Function

    ' 2. Alias変換 (Transaction.TrueCommand の更新)
    Private Function CheckTrueCommand(dat As RpaDataWrapper) As Integer
        dat.Transaction.TrueCommand = vbNullString

        Dim [key] As String = dat.Transaction.CommandText
        Dim [keys] As List(Of String) = dat.Initializer.MyCommandDictionary.Keys.ToList

        ' 初期値
        dat.Transaction.TrueCommand = dat.Transaction.MainCommand

        If [keys].Contains([key]) Then
            If dat.Initializer.MyCommandDictionary([key]).IsEnabled Then
                dat.Transaction.TrueCommand = dat.Initializer.MyCommandDictionary([key]).TrueCommand
            End If
        End If

        ' 自動コマンドの場合、正式名称をセットする
        If dat.Transaction.IsAutoCommand Then
            dat.Transaction.TrueCommand = dat.Transaction.MainCommand
        End If

        'Dim [alias] = dat.Initializer.MyCommandDictionary.Where(
        '    Function(pair)
        '        If pair.Value.Alias = dat.Transaction.MainCommand Then
        '            Return True
        '        Else
        '            Return False
        '        End If
        '    End Function
        ')

        'If Not String.IsNullOrEmpty([alias](0).Key) Then
        '    dat.Transaction.TrueCommand = [alias](0).Key
        'Else
        '    dat.Transaction.TrueCommand = dat.Transaction.MainCommand
        'End If

        Return 0
    End Function

    ' 3. Transaction.UtilityCommand の更新
    ' UtilityCommand の場合、 Transaction.TrueCommand も再更新する
    Private Function CheckUtilityCommand(dat As RpaDataWrapper) As Integer
        dat.Transaction.UtilityCommand = vbNullString

        Dim texts() As String
        texts = dat.Transaction.CommandText.Split(" ")

        If dat.Project IsNot Nothing Then
            If dat.Project.SystemUtilities.Count > 0 Then
                If dat.Project.SystemUtilities.ContainsKey(dat.Transaction.TrueCommand) Then
                    dat.Transaction.UtilityCommand = dat.Transaction.TrueCommand
                    dat.Transaction.TrueCommand = texts(1)
                End If
            End If
        End If

        Return 0
    End Function

    ' 4. Transaction.Parameters, .ParametersText の更新
    Private Function CheckCommandParameters(dat As RpaDataWrapper) As Integer
        dat.Transaction.ParametersText = vbNullString
        dat.Transaction.Parameters = New List(Of String)

        Dim texts() As String
        texts = dat.Transaction.CommandText.Split(" ")

        Dim ex As Integer
        If String.IsNullOrEmpty(dat.Transaction.UtilityCommand) Then
            ex = 0
        Else
            ex = 1
        End If

        Dim cnt As Integer = 0
        For Each p In texts
            If cnt > ex Then
                dat.Transaction.Parameters.Add(p)
                dat.Transaction.ParametersText &= $"{p} "
            End If
            cnt += 1
        Next

        If Not String.IsNullOrEmpty(dat.Transaction.ParametersText) Then
            dat.Transaction.ParametersText _
                = dat.Transaction.ParametersText.TrimEnd(" ")
        End If
        Return 0
    End Function

    Private Function CheckChainBindingCommand(dat As RpaDataWrapper) As Integer
        If Me.ChainBindingCommands.Count > 0 Then
            Return 0
        End If

        Dim cmds As List(Of String) = dat.Transaction.CommandText.Split("&"c).ToList
        If cmds.Count <= 1 Then
            Return 0
        End If

        ' List(Of String).ForEachメソッドが機能しない・・・
        For Each cmd In cmds
            cmd = cmd.Trim
            Me.ChainBindingCommands.Add(cmd)
        Next

        dat.Transaction.CommandText = cmds(0)
        Me.ChainBindingCommands = RpaModule.Pop(Of List(Of String))(Me.ChainBindingCommands)
        Return 0
    End Function

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
                    dat.Transaction.CommandTextLogFileName,
                    dat.Transaction.SystemMonthlyLogFiles(mon),
                    True
                )
                sw = New StreamWriter(
                    dat.Transaction.CommandTextLogFileName,
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

            Dim logdirs() As String = dat.Transaction.SystemMonthlyLogDirectories
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

    Public Async Sub Main(dat As RpaDataWrapper)
        Dim i As Integer = BackupMonthlyLog(dat)               ' 当月ログ保存

        Do
            Call CuiLoop(dat)
        Loop Until dat.Transaction.ExitFlag

        ' プロジェクトセーブ
        '-----------------------------------------------------------------------------------------'
        If dat.Project IsNot Nothing Then
            If File.Exists(dat.Project.SystemJsonChangedFileName) Then
                Dim yorn As String = vbNullString
                Do
                    yorn = vbNullString
                    Console.WriteLine()
                    Console.WriteLine($"Projectに変更があります。変更を保存しますか？ (y/n)")
                    yorn = RpaReadLine(dat, RpaReadLineMode.YesNo, $"Projectに変更があります。変更を保存しますか？ (y/n)")
                Loop Until yorn = "y" Or yorn = "n"
                If yorn = "y" Then
                    RpaModule.Save(dat.Project.SystemJsonFileName, dat.Project, dat.Project.SystemJsonChangedFileName)
                Else
                    File.Delete(dat.Project.SystemJsonChangedFileName)
                End If
            End If
        End If
        '-----------------------------------------------------------------------------------------'

        ' イニシャライザセーブ
        '-----------------------------------------------------------------------------------------'
        If File.Exists(RpaInitializer.SystemIniChangedFileName) Then
            Dim yorn2 As String = vbNullString
            Do
                yorn2 = vbNullString
                Console.WriteLine()
                Console.WriteLine($"Initializerに変更があります。変更を保存しますか？ (y/n)")
                yorn2 = dat.Transaction.ShowRpaIndicator(dat)
            Loop Until yorn2 = "y" Or yorn2 = "n"
            If yorn2 = "y" Then
                RpaModule.Save(RpaCui.SystemIniFileName, dat.Initializer, RpaInitializer.SystemIniChangedFileName)
            Else
                File.Delete(RpaInitializer.SystemIniChangedFileName)
            End If
        End If
        '-----------------------------------------------------------------------------------------'

        Call dat.Transaction.SaveCommandLogs()    ' ログセーブ
        Dim k As Integer = SendMonthlyLog(dat)    ' 当月ログ送信
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

    ' Gui対応の為、Publicプロシージャに変更
    Public Async Sub CuiLoop(dat As RpaDataWrapper)
        Dim cmdlog As RpaTransaction.CommandLogData = Nothing
        Const RESULT_S As String = "Success"
        Const RESULT_F As String = "Failed"
        Const RESULT_E As String = "Exception"
        Dim fastcmdflag As Boolean = False
        Dim latecmdflag As Boolean = False
        Try
            Me.IsRunTime = True
            Dim autocmdflag As Boolean = False
            dat.Transaction.IsAutoCommand = False

            ' 一度のみ実行
            If Not Me.IsUpdatedBindingCommandsCreated Then
                Dim a As Integer = CreateFastBindingCommands(dat)
            End If

            If Me.LateBindingCommands.Count > 0 Then
                dat.Transaction.IsAutoCommand = True
                autocmdflag = True
                latecmdflag = True
                dat.Transaction.CommandText = Me.LateBindingCommands(0)
                Me.LateBindingCommands = RpaModule.Pop(Of List(Of String))(Me.LateBindingCommands)
            ElseIf Me.FastBindingCommands.Count > 0 Then
                Call RpaWriteLine($"更新プログラムを実行中・・・")
                dat.Transaction.IsAutoCommand = True
                autocmdflag = True
                fastcmdflag = True
                If Not IsFastBindingCommandsExecuted Then
                    IsFastBindingCommandsExecuted = True
                End If
                dat.Transaction.CommandText = Me.FastBindingCommands(0)
                Me.FastBindingCommands = RpaModule.Pop(Of List(Of String))(Me.FastBindingCommands)
            ElseIf Me.ChainBindingCommands.Count > 0 Then
                dat.Transaction.CommandText = Me.ChainBindingCommands(0)
                Me.ChainBindingCommands = RpaModule.Pop(Of List(Of String))(Me.ChainBindingCommands)
            ElseIf Me.ExecuteMode = ExecuteModeNumber.RpaGui Then
                dat.Transaction.CommandText = Me.GuiCommandText
            Else
                dat.Transaction.CommandText = dat.Transaction.ShowRpaIndicator(dat)
            End If

            ' コマンドチェーンのチェック
            Dim b As Integer = CheckChainBindingCommand(dat)

            Dim c As Integer = CheckAlias(dat)

            cmdlog = New RpaTransaction.CommandLogData With {
                .UserName = dat.Initializer.UserName,
                .RunDate = (DateTime.Now).ToString,
                .CommandText = dat.Transaction.CommandText,
                .AutoCommandFlag = autocmdflag
            }

            If dat.Project Is Nothing Then
                cmdlog.ProjectArchTypeName = vbNullString
                cmdlog.ProjectName = vbNullString
                cmdlog.RobotName = vbNullString
            Else
                cmdlog.ProjectArchTypeName = dat.Project.SystemArchTypeName
                cmdlog.ProjectName = dat.Project.ProjectName
                cmdlog.RobotName = dat.Project.RobotName
            End If

            dat.Transaction.CommandText _
                = dat.Transaction.CommandText.TrimEnd(" ")

            Dim i As Integer = CheckMainCommand(dat)
            Dim j As Integer = CheckTrueCommand(dat)
            Dim k As Integer = CheckUtilityCommand(dat)
            Dim h As Integer = CheckCommandParameters(dat)
            Dim cmd = CreateCommand(dat)

            cmdlog.UtilityCommand = dat.Transaction.UtilityCommand

            ' コマンドチェック
            '---------------------------------------------------------------------------------'
            Dim err0 As String = $"コマンド  '{dat.Transaction.CommandText}' はありません"
            If cmd Is Nothing Then
                Call RpaWriteLine(err0)
                Call RpaWriteLine()
                cmdlog.Result = RESULT_F
                cmdlog.ResultString = err0
                Throw (New Exception)
            End If

            Dim err1 As String = $"パラメータ数が範囲外です: 許容パラメータ数範囲={cmd.ExecutableParameterCount(0)}~{cmd.ExecutableParameterCount(1)}"
            If cmd.ExecutableParameterCount(0) > dat.Transaction.Parameters.Count Then
                Call RpaWriteLine(err1)
                Call RpaWriteLine()
                cmdlog.Result = RESULT_F
                cmdlog.ResultString = err1
                Throw (New Exception)
            End If
            If cmd.ExecutableParameterCount(1) < dat.Transaction.Parameters.Count Then
                Call RpaWriteLine(err1)
                Call RpaWriteLine()
                cmdlog.Result = RESULT_F
                cmdlog.ResultString = err1
                Throw (New Exception)
            End If

            Dim err2 As String = $"プロジェクトが存在しないため、実行できません"
            If Not cmd.ExecuteIfNoProject Then
                If dat.Project Is Nothing Then
                    Call RpaWriteLine(err2)
                    Call RpaWriteLine()
                    cmdlog.Result = RESULT_F
                    cmdlog.ResultString = err2
                    Throw (New Exception)
                End If
            End If

            Dim err3 As String = $"このプロジェクト構成では、コマンド '{dat.Transaction.CommandText}' は実行できません"
            Dim archs As String() = cmd.ExecutableProjectArchitectures
            If Not archs.Contains("AllArchitectures") Then
                If Not archs.Contains(dat.Project.SystemArchTypeName) Then
                    Call RpaWriteLine(err3)
                    Call RpaWriteLine()
                    cmdlog.Result = RESULT_F
                    cmdlog.ResultString = err3
                    Throw (New Exception)
                End If
            End If

            Dim err4 As String = $"このユーザでは、コマンド '{dat.Transaction.CommandText}' は実行できません"
            Dim users As String() = cmd.ExecutableUser
            If Not users.Contains("AllUser") Then
                If Not users.Contains(dat.Initializer.UserLevel) Then
                    Call RpaWriteLine(err4)
                    Call RpaWriteLine()
                    cmdlog.Result = RESULT_F
                    cmdlog.ResultString = err4
                    Throw (New Exception)
                End If
            End If

            Dim err5 As String = $"コマンドは実行出来ませんでした"
            If Not cmd.CanExecute(dat) Then
                Call RpaWriteLine(err5)
                Call RpaWriteLine()
                cmdlog.Result = RESULT_F
                cmdlog.ResultString = err5
                Throw (New Exception)
            End If
            '---------------------------------------------------------------------------------'

            Dim [from] As DateTime = DateTime.Now
            If Me.ExecuteMode = ExecuteModeNumber.RpaCui Then
                Me.ReturnCode = cmd.Execute(dat)
            ElseIf Me.ExecuteMode = ExecuteModeNumber.RpaGui Then
                Dim t As Task(Of Integer) = Task.Run(
                    Function()
                        Dim x As Integer = cmd.Execute(dat)
                        Return x
                    End Function
                )
                Await t
                Me.ReturnCode = t.Result
            End If
            Dim [to] As DateTime = DateTime.Now
            cmdlog.ExecuteTime = [to] - [from]

            cmdlog.Result = RESULT_S
            cmdlog.ResultString = vbNullString

            If dat.Transaction.ExitFlag Then
                Me.LateExitFlag = True
            End If
            If Me.LateExitFlag And Me.FastBindingCommands.Count = 0 And Me.LateBindingCommands.Count = 0 And Me.ChainBindingCommands.Count = 0 Then
                dat.Transaction.ExitFlag = True
            Else
                dat.Transaction.ExitFlag = False
            End If
        Catch ex As Exception
            ' Exceptionを利用するのは気持ち悪いので、修正するかも・・・
            ' cmdlog.Result が空文字の場合、例外が発生する場合
            If String.IsNullOrEmpty(cmdlog.Result) Then
                Call RpaWriteLine($"({Me.GetType.Name}) {ex.Message}")
                Call RpaWriteLine()
                cmdlog.Result = RESULT_E
                cmdlog.ResultString = ex.Message
            End If

            If latecmdflag Or fastcmdflag Then
                Call RpaWriteLine($"自動生成されたコマンド '{dat.Transaction.CommandText}' は失敗しました")
                Call RpaWriteLine()
                Me.LateBindingCommands = New List(Of String)
                Me.FastBindingCommands = New List(Of String)
            End If

            If fastcmdflag Then
                ' 今のところファストコマンドはアップデート後コマンドしかないので通用するが・・・
                Me.IsUpdatedBindingCommandsFailed = True
            End If
            If Me.LateExitFlag Then
                dat.Transaction.ExitFlag = True
            End If
        Finally
            ' 今のところファストコマンドはアップデート後コマンドしかないので通用するが・・・
            If Me.IsFastBindingCommandsExecuted And Me.FastBindingCommands.Count = 0 And (Not Me.IsUpdatedBindingCommandsFailed) Then
                Dim a As Integer = CompleteUpdate(dat)
            End If

            dat.Transaction.CommandLogs.Add(cmdlog)
            dat.Transaction.Parameters = New List(Of String)
            Me.IsRunTime = False
        End Try
    End Sub

    Public Sub RpaWriteLine(Optional wline As String = vbNullString)
        ' Cuiの場合
        If Me.ExecuteMode = ExecuteModeNumber.RpaCui Then
            Console.WriteLine(wline)
        ElseIf Me.ExecuteMode = ExecuteModeNumber.RpaGui Then
            Me.OutputText &= $"{wline}{vbCrLf}"
        End If
    End Sub

    Public Sub RpaReadLine(ByRef dat As RpaDataWrapper, ByVal mode As Integer, Optional wline As String = vbNullString)
        ' Cuiの場合
        If Me.ExecuteMode = ExecuteModeNumber.RpaCui Then
            If dat.Project IsNot Nothing Then
                Console.Write($"{dat.Project.GuideString}")
            Else
                Console.Write("NoRpa>")
            End If
            Return Console.ReadLine()
        End If
        If Me.ExecuteMode = ExecuteModeNumber.RpaGui Then
            If mode = RpaGuiReadLineMode.YesNo Then
                Return MessageBox.Show(wline, MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2)
            End If
        End If
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
