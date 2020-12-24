Imports System.IO
Public Class RpaSystem

    ' 機能はここに追加
    '---------------------------------------------------------------------------------------------'
    Private ReadOnly Property ExecuteHandler(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Object
        Get
            Dim cmd As Object
            If rpa Is Nothing Then
                ' rpa なしで実行可
                Select Case trn.MainCommand
                    Case "setup" : cmd = New SetupRpaCommand
                    Case "exit" : cmd = New ExitCommand
                    Case Else : cmd = Nothing
                End Select
            Else
                ' rpa ありで実行可
                Select Case trn.MainCommand
                    Case "setup" : cmd = New SetupRpaCommand
                    Case "save" : cmd = New SaveCommand
                    Case "load" : cmd = New LoadCommand
                    Case "run" : cmd = New RunRobotCommand
                    Case "clone" : cmd = New CloneProjectCommand
                    Case "select" : cmd = New SelectProjectCommand
                    Case "exit" : cmd = New ExitCommand
                    Case "addutility" : cmd = New AddUtilityCommand
                    'Case "project" : cmd = AddressOf ShowCurrentProject
                    'Case "savejson" : cmd = AddressOf SaveJson
                    'Case "update" : cmd = AddressOf UpdateProject
                    'Case "show" : cmd = AddressOf ShowProject
                    Case Else : cmd = Nothing
                End Select

                If cmd Is Nothing And rpa.SystemUtilities.Count > 0 Then
                    For Each util In rpa.SystemUtilities
                        cmd = util.Value.UtilityObject.ExecuteHandler(trn, rpa)
                        If cmd IsNot Nothing Then
                            Exit For
                        End If
                    Next
                End If
            End If
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

    ' システムのロード
    '---------------------------------------------------------------------------------------------'
    Private Function LoadSystem(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
        Dim [load] = rpa.Load(rpa.SystemJsonFileName)
        Dim i = 0
        If [load] Is Nothing Then
            Console.WriteLine($"JsonFile '{rpa.SystemJsonFileName}' のロードに失敗しました。")
            Return 1000
        End If
        If Not Directory.Exists([load].RootDirectory) Then
            Console.WriteLine($"RootDirectory '{[load].RootDirectory}' がありません")
            Console.WriteLine($"ファイル '{[load].SystemJsonFileName}' の 'RootDirectory' に任意のパスを書いてください")
            Console.WriteLine("ファイルを保存した後、アプリケーションを再起動してください")
            Console.ReadLine()
            trn.ExitFlag = True
            Return 1000
        End If
        If Not Directory.Exists([load].MyDirectory) Then
            Console.WriteLine($"MyDirectory '{[load].MyDirectory}' がありません")
            Console.WriteLine($"ファイル '{[load].SystemJsonFileName}' の 'MyDirectory' に任意のパスを書いてください")
            Console.WriteLine("ファイルを保存した後、アプリケーションを再起動してください")
            Console.ReadLine()
            trn.ExitFlag = True
            Return 1000
        End If
        rpa = [load]
        Console.WriteLine($"JsonFile '{rpa.SystemJsonFileName}' をロードしました。")
        Return 0
    End Function
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

    ' ユーティリティの追加
    '---------------------------------------------------------------------------------------------'
    Private Function AddUtility(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
        Dim uobj As RpaUtility
        Dim obj As Object
        Dim util As String
        If trn.Parameters.Count = 0 Then
            Console.WriteLine("ユーティリティが指定されていません")
            Return 1000
        End If
        util = trn.Parameters(0)
        obj = RpaCodes.RpaUtility(util)
        If obj Is Nothing Then
            Console.WriteLine($"指定のユーティリティ '{util}' は存在しません")
            Return 1000
        End If
        If rpa.SystemUtilities.ContainsKey(util) Then
            Console.WriteLine($"指定のユーティリティ '{util}' は既に追加されています")
            Return 1000
        End If
        uobj = New RpaUtility With {
            .UtilityName = util,
            .UtilityObject = obj
        }
        rpa.SystemUtilities.Add(util, uobj)
        Console.WriteLine($"指定のユーティリティ '{util}' を追加しました")
        Return 0
    End Function
    '---------------------------------------------------------------------------------------------'

    ' Ｊｓｏｎのセーブ
    '---------------------------------------------------------------------------------------------'
    Private Delegate Sub SaveJsonDelegater(ByVal json As String, ByVal obj As Object)
    Private Function SaveJson(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
        Dim act As Action(Of String, Object)
        act = Sub(json, obj)
                  Call obj.Save(json, obj)
                  obj = obj.Load(json)
                  If obj IsNot Nothing Then
                      Console.WriteLine($"'{json}' のセーブに成功しました")
                      obj = obj
                  End If
              End Sub
        If trn.Parameters.Count = 0 Then
            Call act(rpa.RootProjectJsonFileName, rpa.RootProjectObject)
            Call act(rpa.MyProjectJsonFileName, rpa.MyProjectObject)
        Else
            If trn.Parameters(0) = "root" Then
                Call act(rpa.RootProjectJsonFileName, rpa.RootProjectObject)
            ElseIf trn.Parameters(0) = "myproject" Then
                Call act(rpa.MyProjectJsonFileName, rpa.MyProjectObject)
            End If
        End If
        Return 0
    End Function

    Private Sub _SaveAndCheckJson(ByVal file As String, ByVal obj As Object)
    End Sub
    '---------------------------------------------------------------------------------------------'

    ' マクロの更新
    ''---------------------------------------------------------------------------------------------'
    'Private Function UpdateMacro(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
    '    Dim bas As String
    '    If trn.Parameters.Count = 0 Then
    '        Console.WriteLine("パラメータが指定されていません: " & trn.CommandText)
    '        Return 1000
    '    End If

    '    For Each p In trn.Parameters
    '        bas = RpaProject.SYSTEM_SCRIPT_DIRECTORY & "\" & p
    '        If File.Exists(bas) Then
    '            Console.WriteLine($"指定マクロ '{p}' をインストールします")
    '            Call rpa.InvokeMacro("MacroImporter.Main", {bas})
    '        Else
    '            Console.WriteLine($"指定マクロ '{p}' は存在しません")
    '        End If
    '    Next
    '    Return 0
    'End Function

    '---------------------------------------------------------------------------------------------'

    ' ロボットの起動
    '---------------------------------------------------------------------------------------------'
    'Private Function RunRobot(ByRef trn As RpaTransaction, ByRef rpa As IntranetClientServerProject) As Integer
    '    Dim i = 0
    '    Call rpa.MyProjectObject.SetData(trn, rpa)
    '    If rpa.MyProjectObject.CanExecute() Then
    '        i = rpa.MyProjectObject.Execute()
    '    Else
    '        Console.WriteLine("ロボットの起動条件を満たしていません")
    '        i = 1000
    '    End If
    '    Return i
    'End Function
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

    Public Sub Main(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer)
        Dim cmd = ExecuteHandler(trn, rpa, ini)
        If cmd IsNot Nothing Then
            trn.ReturnCode = cmd.Execute(trn, rpa, ini)
        Else
            Console.WriteLine($"コマンド : '{trn.CommandText}' はありません")
        End If
    End Sub
End Class
