Imports System.IO

Public Class RpaSystem
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
                Me._CommandDictionary.Add("updaterobot", (New UpdateRobotCommand))
                Me._CommandDictionary.Add("exportmyrobotjson", (New ExportMyRobotJsonCommand))
                Me._CommandDictionary.Add("exportrootrobotjson", (New ExportRootRobotJsonCommand))
                Me._CommandDictionary.Add("removemycommand", (New RemoveMyCommandCommand))
                Me._CommandDictionary.Add("activatemycommand", (New ActivateMyCommandCommand))
                Me._CommandDictionary.Add("showmyrobotjson", (New ShowMyRobotJsonCommand))
                Me._CommandDictionary.Add("showmyrobotproperties", (New ShowMyRobotPropertiesCommand))
                Me._CommandDictionary.Add("pushrobot", (New PushRobotCommand))
                Me._CommandDictionary.Add("showrobotreadme", (New ShowRobotReadMeCommand))
            End If
            Return Me._CommandDictionary
        End Get
    End Property

    ' 機能はここに追加
    ' ここでインスタンス化されたコマンドには、RpaDataWrapper型を持たせることが出来る
    ' (おそらく、コンパイル時に型の情報を検査するのだと思われる)
    '---------------------------------------------------------------------------------------------'
    Public ReadOnly Property CommandHandler(dat As RpaDataWrapper) As Object
        Get
            ' Aliasの変換
            Dim [alias] = dat.Initializer.MyCommandDictionary.Where(
                Function(pair)
                    If pair.Value.Alias = dat.Transaction.MainCommand Then
                        Return True
                    Else
                        Return False
                    End If
                End Function
            )
            If Not String.IsNullOrEmpty([alias](0).Key) Then
                dat.Transaction.MainCommand = [alias](0).Key
            End If


            ' コマンド生成
            Dim cmdtxt As String = dat.Transaction.MainCommand
            Dim cmd = Nothing
            If Me.CommandDictionary.ContainsKey(cmdtxt) Then
                cmd = Me.CommandDictionary(cmdtxt)
            Else
                cmd = Nothing
            End If


            ' ユーティリティコマンド
            If dat.Project IsNot Nothing Then
                If (dat.Project.SystemUtilities.Count > 0) And (cmd Is Nothing) Then
                    If dat.Project.SystemUtilities.ContainsKey(dat.Transaction.MainCommand) Then
                        dat.Transaction.MainCommand = dat.Transaction.Parameters(0)
                        dat.Transaction.Parameters = RpaModule.Pop(Of List(Of String))(dat.Transaction.Parameters)
                        cmd = dat.Project.SystemUtilities(cmdtxt).UtilityObject.CommandHandler(dat)
                    End If
                End If
            End If


            ' コマンド無効化
            If dat.Initializer.MyCommandDictionary.ContainsKey(cmdtxt) Then
                If Not dat.Initializer.MyCommandDictionary(cmdtxt).IsEnabled Then
                    cmd = Nothing
                End If
            End If

            Return cmd
        End Get
    End Property
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

    Public Sub Main(ByRef dat As RpaDataWrapper)
        Dim i As Integer = BackupMonthlyLog(dat)  ' 当月ログ保存

        Do
            Call CuiLoop(dat)
        Loop Until dat.Transaction.ExitFlag

        '' アップデートチェック
        ''-----------------------------------------------------------------------------------------'
        'Dim j As Integer
        'If dat.Project.SystemArchTypeName = (New IntranetClientServerProject).GetType.Name Then
        '    j = UpdateCheck(dat)
        'End If
        ''-----------------------------------------------------------------------------------------'

        ' プロジェクトセーブ
        '-----------------------------------------------------------------------------------------'
        If dat.Project IsNot Nothing Then
            If File.Exists(dat.Project.SystemJsonChangedFileName) Then
                Dim yorn As String = vbNullString
                Do
                    yorn = vbNullString
                    Console.WriteLine()
                    Console.WriteLine($"Projectに変更があります。変更を保存しますか？ (y/n)")
                    yorn = dat.Transaction.ShowRpaIndicator(dat)
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

    ' IntranetClientServerProjectで実行
    'Private Function UpdateCheck(ByRef dat As RpaDataWrapper) As Integer
    '    If Not dat.Project.IsUpdateAvailable Then
    '        Return 0
    '    End If

    '    Dim idx As String = vbNullString
    '    If dat.Project.AutoUpdate Then
    '        idx = "0"
    '    Else
    '        Do
    '            idx = vbNullString
    '            Console.WriteLine()
    '            Console.WriteLine($"0 ...更新して終了")
    '            Console.WriteLine($"1 ...終了")
    '            idx = dat.Transaction.ShowRpaIndicator(dat)
    '            If idx = "0" Or idx = "1" Then
    '                Exit Do
    '            End If
    '        Loop Until False
    '    End If

    '    If idx = "1" Then
    '        Return 0
    '    End If

    '    Dim jh As New RpaCui.JsonHandler(Of List(Of RpaUpdater))
    '    Dim rrus As List(Of RpaUpdater) = jh.Load(Of List(Of RpaUpdater))(dat.Project.RootRobotsUpdateFile)
    '    Dim uru As RpaUpdater
    '    rrus.Sort(
    '        Function(before, after)
    '            Return (before.ReleaseDate < after.ReleaseDate)
    '        End Function
    '    )
    '    uru = rrus.Last

    '    ' リテラル部分は後で修正する（面倒なので、リテラルで仮置きしている）
    '    '-----------------------------------------------------------------------------------------'
    '    dat.Transaction.AutoCommandText.Add($"updaterobot {uru.ReleaseId}")
    '    '-----------------------------------------------------------------------------------------'

    '    Return 0
    'End Function

    Private Sub CuiLoop(ByRef dat As RpaDataWrapper)
        Const RESULT_S As String = "Success"
        Const RESULT_F As String = "Failed"
        Const RESULT_E As String = "Exception"

        Dim cmdlog As RpaTransaction.CommandLogData = Nothing
        Dim autocmdflag As Boolean = False

        Try
            dat.Transaction.CommandText = dat.Transaction.ShowRpaIndicator(dat)
            Do
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

                Call dat.Transaction.CreateCommand()

                Dim cmd = CommandHandler(dat)

                Dim err0 As String = $"コマンド  '{dat.Transaction.CommandText}' はありません"
                If cmd Is Nothing Then
                    Console.WriteLine(err0)
                    Console.WriteLine()
                    cmdlog.Result = RESULT_F
                    cmdlog.ResultString = err0
                    Throw (New Exception)
                End If

                Dim err1 As String = $"パラメータ数が範囲外です: 許容パラメータ数範囲={cmd.ExecutableParameterCount(0)}~{cmd.ExecutableParameterCount(1)}"
                If cmd.ExecutableParameterCount(0) > dat.Transaction.Parameters.Count Then
                    Console.WriteLine(err1)
                    Console.WriteLine()
                    cmdlog.Result = RESULT_F
                    cmdlog.ResultString = err1
                    Throw (New Exception)
                End If
                If cmd.ExecutableParameterCount(1) < dat.Transaction.Parameters.Count Then
                    Console.WriteLine(err1)
                    Console.WriteLine()
                    cmdlog.Result = RESULT_F
                    cmdlog.ResultString = err1
                    Throw (New Exception)
                End If

                Dim err2 As String = $"プロジェクトが存在しないため、実行できません"
                If Not cmd.ExecuteIfNoProject Then
                    If dat.Project Is Nothing Then
                        Console.WriteLine(err2)
                        Console.WriteLine()
                        cmdlog.Result = RESULT_F
                        cmdlog.ResultString = err2
                        Throw (New Exception)
                    End If
                End If


                Dim err3 As String = $"このプロジェクト構成では、コマンド '{dat.Transaction.CommandText}' は実行できません"
                Dim archs As String() = cmd.ExecutableProjectArchitectures
                If Not archs.Contains("AllArchitectures") Then
                    If Not archs.Contains(dat.Project.SystemArchTypeName) Then
                        Console.WriteLine(err3)
                        Console.WriteLine()
                        cmdlog.Result = RESULT_F
                        cmdlog.ResultString = err3
                        Throw (New Exception)
                    End If
                End If

                Dim err4 As String = $"このユーザでは、コマンド '{dat.Transaction.CommandText}' は実行できません"
                Dim users As String() = cmd.ExecutableUser
                If Not users.Contains("AllUser") Then
                    If Not users.Contains(dat.Initializer.UserLevel) Then
                        Console.WriteLine(err4)
                        Console.WriteLine()
                        cmdlog.Result = RESULT_F
                        cmdlog.ResultString = err4
                        Throw (New Exception)
                    End If
                End If

                Dim err5 As String = $"コマンドは実行出来ませんでした"
                If Not cmd.CanExecute(dat) Then
                    Console.WriteLine(err5)
                    Console.WriteLine()
                    cmdlog.Result = RESULT_F
                    cmdlog.ResultString = err5
                    Throw (New Exception)
                End If

                Dim [from] As DateTime = DateTime.Now
                dat.Transaction.ReturnCode = cmd.Execute(dat)
                Dim [to] As DateTime = DateTime.Now
                cmdlog.ExecuteTime = [to] - [from]

                cmdlog.Result = RESULT_S
                cmdlog.ResultString = vbNullString
                dat.Transaction.CommandLogs.Add(cmdlog)

                ' 自動生成コマンド
                '---------------------------------------------------------------------------------'
                If dat.Transaction.AutoCommandText.Count = 0 Then
                    Exit Do
                End If

                autocmdflag = True
                dat.Transaction.CommandText = dat.Transaction.AutoCommandText(0)
                dat.Transaction.AutoCommandText = RpaModule.Pop(Of List(Of String))(dat.Transaction.AutoCommandText)
                '---------------------------------------------------------------------------------'
            Loop Until False
        Catch ex As Exception
            ' Exceptionを利用するのは気持ち悪いので、修正するかも・・・
            ' cmdlog.Result が空文字の場合、例外が発生する場合
            If String.IsNullOrEmpty(cmdlog.Result) Then
                Console.WriteLine($"({Me.GetType.Name}) {ex.Message}")
                Console.WriteLine()
                cmdlog.Result = RESULT_E
                cmdlog.ResultString = ex.Message
            End If
            dat.Transaction.CommandLogs.Add(cmdlog)
        Finally
            dat.Transaction.Parameters = New List(Of String)
        End Try
    End Sub
End Class
