Imports System.IO

Public Class RpaSystem
    ' 機能はここに追加
    ' ここでインスタンス化されたコマンドには、RpaDataWrapper型を持たせることが出来る
    ' (おそらく、コンパイル時に型の情報を検査するのだと思われる)
    '---------------------------------------------------------------------------------------------'
    Public ReadOnly Property CommandHandler(dat As RpaDataWrapper) As Object
        Get
            Dim cmd = Nothing

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

            Dim cmdtxt As String = dat.Transaction.MainCommand

            Select Case cmdtxt
                    'Case "setup" : cmd = New SetupRpaCommand
                Case "load" : cmd = New LoadCommand
                Case "exit" : cmd = New ExitCommand
                Case "showmycommands" : cmd = New ShowMyCommandCommand
                Case "exportmycommands" : cmd = New ExportMyCommandsCommand
                Case "importmycommands" : cmd = New ImportMyCommandsCommand
                Case "changecommandenabled" : cmd = New ChangeCommandEnabledCommand
                Case "setup" : cmd = New SetupRpaCommand
                Case "saveproject" : cmd = New SaveProjectCommand
                Case "saveinitializer" : cmd = New SaveInitializerCommand
                Case "load" : cmd = New LoadCommand
                Case "runrobot" : cmd = New RunRobotCommand
                Case "clonerobot" : cmd = New CloneRobotCommand
                Case "attachrobot" : cmd = New AttachRobotCommand
                Case "addutility" : cmd = New AddUtilityCommand
                Case "addcommandalias" : cmd = New AddCommandAliasCommand
                Case "addrobotalias" : cmd = New AddRobotAliasCommand
                Case "removerobot" : cmd = New RemoveRobotCommand
                Case "showmycommands" : cmd = New ShowMyCommandCommand
                Case "exportmycommands" : cmd = New ExportMyCommandsCommand
                Case "importmycommands" : cmd = New ImportMyCommandsCommand
                Case "changecommandenabled" : cmd = New ChangeCommandEnabledCommand
                    'Case "project" : cmd = AddressOf ShowCurrentProject
                    'Case "savejson" : cmd = AddressOf SaveJson
                    'Case "update" : cmd = AddressOf UpdateProject
                    'Case "show" : cmd = AddressOf ShowProject
                Case "removeproject" : cmd = New RemoveProjectCommand
                Case "newproject" : cmd = New NewProjectCommand
                Case "changeprojectproperty" : cmd = New ChangeProjectPropertyCommand
                Case "showprojectproperties" : cmd = New ShowProjectPropertiesCommand
                Case "changeprojectpropertyusingfolderbrowser" : cmd = New ChangeProjectPropertyUsingFolderBrowserCommand
                Case "showinitializerproperties" : cmd = New ShowInitializerPropertiesCommand
                Case "changeinitializerproperty" : cmd = New ChangeInitializerPropertyCommand
                Case Else : cmd = Nothing
            End Select

            ' コマンド無効化
            If dat.Initializer.MyCommandDictionary.ContainsKey(cmdtxt) Then
                If Not dat.Initializer.MyCommandDictionary(cmdtxt).IsEnabled Then
                    cmd = Nothing
                End If
            End If
            'If dat.Transaction.Modes.Count > 0 Then
            '    Select Case dat.Transaction.Modes.Last
            '        Case (New SetupRpaCommand).GetType.Name
            '            Select Case dat.Transaction.MainCommand
            '                Case "newproject" : cmd = New NewProjectCommand
            '                Case "setupproject" : cmd = New SetupProjectCommand
            '                Case "setupinitializer" : cmd = New SetupInitializerCommand
            '                Case "removeproject" : cmd = New RemoveProjectCommand
            '                Case "exit" : cmd = New ExitCommand
            '                Case Else : cmd = Nothing
            '            End Select
            '        Case (New SetupProjectCommand).GetType.Name
            '            Select Case dat.Transaction.MainCommand
            '                Case "changeprojectproperty" : cmd = New ChangeProjectPropertyCommand
            '                Case "showprojectproperties" : cmd = New ShowProjectPropertiesCommand
            '                Case "changeprojectpropertyusingfolderbrowser" : cmd = New ChangeProjectPropertyUsingFolderBrowserCommand
            '                Case "exit" : cmd = New ExitCommand
            '                Case Else : cmd = Nothing
            '            End Select
            '        Case (New SetupInitializerCommand).GetType.Name
            '            Select Case dat.Transaction.MainCommand
            '                Case "changeinitializerproperty" : cmd = New ChangeInitializerCommand
            '                Case "showinitializerproperties" : cmd = New ShowInitializerPropertiesCommand
            '                Case "exit" : cmd = New ExitCommand
            '                Case Else : cmd = Nothing
            '            End Select
            '        Case Else : cmd = Nothing
            '    End Select
            'End If

            Return cmd
        End Get
    End Property
    '---------------------------------------------------------------------------------------------'

    ' システムのロード
    '---------------------------------------------------------------------------------------------'
    'Private Function ShowCurrentProject(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
    '    Console.WriteLine($"ProjectName  : {rpa.ProjectName}")
    '    Console.WriteLine($"ProjectAlias : {rpa.ProjectAlias}")
    '    Console.WriteLine($"RootProject  : {rpa.RootProjectDirectory}")
    '    Console.WriteLine($"MyProject    : {rpa.MyProjectDirectory}")
    '    Return 0
    'End Function
    '---------------------------------------------------------------------------------------------'

    ' システムのセーブ
    '---------------------------------------------------------------------------------------------'
    Private Function SaveSystem(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
        Call rpa.Save(rpa.SystemJsonFileName, rpa)
        Console.WriteLine($"JsonFile '{rpa.SystemJsonFileName}' をセーブしました。")
        Return 0
    End Function
    '---------------------------------------------------------------------------------------------'

    ' プロジェクトの確認
    '---------------------------------------------------------------------------------------------'
    Private Function ShowProject(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
        If trn.Parameters.Count = 0 Then
            Console.WriteLine("パラメータが指定されていません")
            Return 1000
        End If
        If trn.Parameters.Last = "Utilities" Then
            Return _ShowProjectUtilities(trn, rpa)
        End If
        Return 1000
    End Function

    ' ユーティリティの確認
    Private Function _ShowProjectUtilities(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
        Console.WriteLine(Strings.StrDup(100, "-"))
        If rpa.SystemUtilities.Count > 0 Then
            For Each util In rpa.SystemUtilities
                Console.WriteLine($"{util.Key}")
            Next
        Else
            Console.WriteLine("ユーティリティなし")
        End If
        Console.WriteLine(Strings.StrDup(100, "-"))
        Return 0
    End Function
    '---------------------------------------------------------------------------------------------'

    ' Ｊｓｏｎのセーブ
    '---------------------------------------------------------------------------------------------'
    'Private Delegate Sub SaveJsonDelegater(ByVal json As String, ByVal obj As Object)
    'Private Function SaveJson(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
    '    Dim act As Action(Of String, Object)
    '    act = Sub(json, obj)
    '              Call obj.Save(json, obj)
    '              obj = obj.Load(json)
    '              If obj IsNot Nothing Then
    '                  Console.WriteLine($"'{json}' のセーブに成功しました")
    '                  obj = obj
    '              End If
    '          End Sub
    '    If trn.Parameters.Count = 0 Then
    '        Call act(rpa.RootProjectJsonFileName, rpa.RootProjectObject)
    '        Call act(rpa.MyProjectJsonFileName, rpa.MyProjectObject)
    '    Else
    '        If trn.Parameters(0) = "root" Then
    '            Call act(rpa.RootProjectJsonFileName, rpa.RootProjectObject)
    '        ElseIf trn.Parameters(0) = "myproject" Then
    '            Call act(rpa.MyProjectJsonFileName, rpa.MyProjectObject)
    '        End If
    '    End If
    '    Return 0
    'End Function

    Private Sub _SaveAndCheckJson(ByVal file As String, ByVal obj As Object)
    End Sub
    '---------------------------------------------------------------------------------------------'


    ' プロジェクトのアップデート
    '---------------------------------------------------------------------------------------------'
    Private Function UpdateProject(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
        'Dim i = -1
        'Dim answer = vbNullString
        'Dim pack As RpaPackage = Nothing

        'If trn.Parameters.Count = 0 Then
        '    Console.WriteLine("パラメータが指定されていません: " & trn.CommandText)
        '    Return 1000
        'End If

        'If Not rpa.IsRootProjectExists Then
        '    Console.WriteLine("プロジェクト: '" & rpa.ProjectName & "' のルートプロジェクトが存在しません")
        '    Console.WriteLine("ダウンロード先が存在しません")
        '    Return 1000
        'End If

        'If rpa.RootProjectUpdatePackages.Count = 0 Then
        '    Console.WriteLine("'" & rpa.RootProjectUpdateDirectory & "' にパッケージリストが存在しません")
        '    Return 1000
        'End If

        'Console.WriteLine("パッケージを検索しています...")
        'For Each p In rpa.RootProjectUpdatePackages
        '    If p.Name = trn.Parameters(0) Then
        '        pack = p
        '    End If
        'Next
        'If pack Is Nothing Then
        '    Console.WriteLine("指定のパッケージ '" & trn.Parameters(0) & "' はパッケージリストに登録されていません")
        '    Return 1000
        'End If

        'If pack.Latest Then
        '    answer = "y"
        'Else
        '    Console.WriteLine("指定のパッケージ '" & pack.Name & "' は最新のパッケージではありません。")
        '    Console.WriteLine("既に前方互換性を持つパッケージをダウンロードしているか確認をしてください")
        '    Console.WriteLine(vbNullString)
        '    Console.WriteLine("ダウンロード済みパッケージ")
        '    Console.WriteLine("-----------------------------------------------------------")
        '    For Each p In rpa.MyProjectUpdatedPackages
        '        Console.WriteLine(p.Name)
        '    Next
        '    If rpa.MyProjectUpdatedPackages.Count = 0 Then
        '        Console.WriteLine("ダウンロード済みパッケージ　なし")
        '    End If
        '    Console.WriteLine("パッケージをダウンロードしますか？ (y/n)")
        '    Do
        '        Console.ReadLine()
        '    Loop Until answer = "y" Or "n"
        'End If

        'If answer = "n" Then
        '    Return 0
        'End If

        'Dim src = vbNullString
        'Dim dsd = vbNullString
        'Dim dst = vbNullString
        'For Each ui In pack.UpdateInfos
        '    src = ui.SourceFile
        '    dsd = ui.DistinationSubDirectory
        '    If Not String.IsNullOrEmpty(dsd) Then
        '        dst = rpa.MyProjectDirectory & "\" & ui.DistinationSubDirectory & "\" & Path.GetFileName(src)
        '    Else
        '        dst = vbNullString
        '    End If

        '    If Not File.Exists(ui.SourceFile) Then
        '        Console.WriteLine("ソースファイル: " & src & " がありません")
        '        Console.WriteLine("詳しくは作成者に問い合わせてください... Author: " & ui.Author)
        '        Continue For
        '    End If

        '    If Not Directory.Exists(dsd) Then
        '        Console.WriteLine("アップデート先ディレクトリ: " & dsd & " がありません")
        '        Continue For
        '    End If

        '    File.Copy(src, dst, True)
        'Next

        Return 0
    End Function
    '---------------------------------------------------------------------------------------------'

    ' プロジェクトのダウンロード
    '---------------------------------------------------------------------------------------------'
    'Private Function CopyProject(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
    '    If trn.Parameters.Count > 0 Then
    '        Call SelectedCopy(trn, rpa)
    '    Else
    '        Call AllCopy(trn, rpa, rpa.RootProjectDirectory, rpa.MyProjectDirectory)
    '    End If
    '    Console.WriteLine("コピーが完了しました")

    '    Return 0
    'End Function

    'Private Sub SelectedCopy(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject)
    '    Dim src = vbNullString
    '    Dim dst = vbNullString
    '    Dim dstp = vbNullString
    '    For Each p In trn.Parameters
    '        src = rpa.RootProjectDirectory & "\" & p
    '        dst = rpa.MyProjectDirectory & "\" & p
    '        dstp = Path.GetDirectoryName(dst)
    '        If Not Directory.Exists(dstp) Then
    '            Console.WriteLine("ディレクトリが存在しません: " & dstp)
    '            Continue For
    '        End If
    '        If File.Exists(src) Then
    '            File.Copy(src, dst, True)
    '            Console.WriteLine($"Directory.CreateDirectory... src: {src}")
    '            Console.WriteLine($"                          => dst: {dst}")
    '        ElseIf Directory.Exists(src) Then
    '            If Directory.Exists(dst) Then
    '                Console.WriteLine("既にディレクトリが存在します: " & dst)
    '            Else
    '                Directory.CreateDirectory(dst)
    '                Console.WriteLine($"Directory.CreateDirectory... src: {src}")
    '                Console.WriteLine($"                          => dst: {dst}")
    '            End If
    '        End If
    '    Next
    'End Sub

    'Private Sub AllCopy(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject,
    '                              ByVal src As String, ByVal dst As String)
    '    If Not Directory.Exists(dst) Then
    '        Directory.CreateDirectory(dst)
    '        Console.WriteLine($"Directory.CreateDirectory... src: {src}")
    '        Console.WriteLine($"                          => dst: {dst}")
    '    End If

    '    Dim fsis As FileSystemInfo()
    '    Dim sdi = New DirectoryInfo(src)
    '    Dim ddi = New DirectoryInfo(dst)
    '    Dim src2 = vbNullString
    '    Dim dst2 = vbNullString
    '    fsis = sdi.GetFileSystemInfos
    '    For Each fsi In fsis
    '        src2 = fsi.FullName
    '        dst2 = ddi.FullName & "\" & fsi.Name
    '        If Not rpa.RootProjectIgnoreList.Contains(src2) Then
    '            If Not rpa.MyProjectIgnoreList.Contains(src2) Then
    '                If ((fsi.Attributes And FileAttributes.Directory) = FileAttributes.Directory) Then
    '                    Call AllCopy(trn, rpa, src2, dst2)
    '                Else
    '                    File.Copy(src2, dst2, True)
    '                    Console.WriteLine("File.Copy... src: " & src2)
    '                    Console.WriteLine("          => dst: " & dst2)
    '                End If
    '            End If
    '        End If
    '    Next
    'End Sub
    '---------------------------------------------------------------------------------------------'


    ' プロジェクトの切り替え
    '---------------------------------------------------------------------------------------------'
    'Private Function SelectProject(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
    '    If trn.Parameters.Count = 0 Then
    '        Console.WriteLine($"プロジェクトが選択されていません")
    '        Return 1000
    '    Else
    '        rpa = New IntranetClientServerProject With {.ProjectName = trn.Parameters(0)}
    '        Return 0
    '    End If
    'End Function
    '---------------------------------------------------------------------------------------------'


    ' プロジェクト終了
    '---------------------------------------------------------------------------------------------'
    'Private Function RpaExit(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
    '    trn.ExitFlag = True
    '    Return 0
    'End Function
    '---------------------------------------------------------------------------------------------'

    Public Sub Main(ByRef dat As RpaDataWrapper)
        dat.Transaction.CommandText = dat.Transaction.ShowRpaIndicator(dat)
        Call dat.Transaction.CreateCommand()

        Dim cmd = CommandHandler(dat)

        If dat.Project IsNot Nothing Then
            If (dat.Project.SystemUtilities.Count > 0) And (cmd Is Nothing) Then
                For Each util In dat.Project.SystemUtilities
                    cmd = util.Value.UtilityObject.ExecuteHandler(dat)
                Next
            End If
        End If

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
            Console.WriteLine(err2)
            Console.WriteLine()
            Exit Sub
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

        Dim err4 As String = $"コマンドは実行出来ませんでした"
        If Not cmd.CanExecute(dat) Then
            Console.WriteLine(err4)
            Console.WriteLine()
            Exit Sub
        End If

        dat.Transaction.ReturnCode = cmd.Execute(dat)
    End Sub
End Class
