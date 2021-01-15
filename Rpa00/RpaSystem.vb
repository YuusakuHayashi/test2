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

    Public Sub Main(ByRef dat As RpaDataWrapper)
        Try
            dat.Transaction.CommandText = dat.Transaction.ShowRpaIndicator(dat)
            Call dat.Transaction.CreateCommand()

            Dim cmd = CommandHandler(dat)

            If cmd Is Nothing Then
                Console.WriteLine($"コマンド : '{dat.Transaction.CommandText}' はありません")
                Console.WriteLine()
                Exit Sub
            End If

            Dim err1 As String = $"パラメータ数が範囲外です: 許容パラメータ数範囲={cmd.ExecutableParameterCount(0)}~{cmd.ExecutableParameterCount(1)}"
            If cmd.ExecutableParameterCount(0) > dat.Transaction.Parameters.Count Then
                Console.WriteLine(err1)
                Console.WriteLine()
                Exit Sub
            End If
            If cmd.ExecutableParameterCount(1) < dat.Transaction.Parameters.Count Then
                Console.WriteLine(err1)
                Console.WriteLine()
                Exit Sub
            End If

            Dim err2 As String = $"プロジェクトが存在しないため、実行できません"
            If Not cmd.ExecuteIfNoProject Then
                If dat.Project Is Nothing Then
                    Console.WriteLine(err2)
                    Console.WriteLine()
                    Exit Sub
                End If
            End If


            Dim err3 As String = $"このプロジェクト構成では、コマンド '{dat.Transaction.CommandText}' は実行できません"
            Dim archs As String() = cmd.ExecutableProjectArchitectures
            If Not archs.Contains("AllArchitectures") Then
                If Not archs.Contains(dat.Project.SystemArchTypeName) Then
                    Console.WriteLine(err3)
                    Console.WriteLine()
                    Exit Sub
                End If
            End If

            Dim err4 As String = $"このユーザでは、コマンド '{dat.Transaction.CommandText}' は実行できません"
            Dim users As String() = cmd.ExecutableUser
            If Not users.Contains("AllUser") Then
                If Not users.Contains(dat.Initializer.UserLevel) Then
                    Console.WriteLine(err4)
                    Console.WriteLine()
                    Exit Sub
                End If
            End If

            Dim err5 As String = $"コマンドは実行出来ませんでした"
            If Not cmd.CanExecute(dat) Then
                Console.WriteLine(err5)
                Console.WriteLine()
                Exit Sub
            End If

            dat.Transaction.ReturnCode = cmd.Execute(dat)

        Catch ex As Exception
            Console.WriteLine($"({Me.GetType.Name}) {ex.Message}")
            Console.WriteLine()
        Finally
            dat.Transaction.Parameters = New List(Of String)
        End Try
    End Sub
End Class
