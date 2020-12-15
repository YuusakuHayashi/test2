Imports System.IO
Public Class RpaSystem
    Private Delegate Function ExecuteDelegater(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer

    ' 機能はここに追加
    '---------------------------------------------------------------------------------------------'
    Private ReadOnly Property ExecuteHandler(ByVal cmd As String) As ExecuteDelegater
        Get
            Dim dlg As ExecuteDelegater
            Select Case cmd
                Case "SaveJson"
                    dlg = AddressOf SaveJson
                'Case "MacroUpdate"
                '    dlg = AddressOf UpdateMacro
                Case "Run"
                    dlg = AddressOf RunRobot
                Case "Update"
                    dlg = AddressOf UpdateProject
                Case "CopyProject"
                    dlg = AddressOf DownloadProject
                Case "Project"
                    dlg = AddressOf SetProject
                Case "AddUtility"
                    dlg = AddressOf AddUtility
                Case "Exit"
                    dlg = AddressOf RpaExit
                Case Else
                    dlg = Nothing
            End Select

            If dlg IsNot Nothing Then
                Return dlg
            End If
        End Get
    End Property
    '---------------------------------------------------------------------------------------------'

    ' ユーティリティの追加
    '---------------------------------------------------------------------------------------------'
    Private Function AddUtility(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
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

        rpa.SystemUtilities.Add(util, obj)
        Return 0
    End Function
    '---------------------------------------------------------------------------------------------'

    ' Ｊｓｏｎのセーブ
    '---------------------------------------------------------------------------------------------'
    Private Delegate Sub SaveJsonDelegater(ByVal json As String, ByVal obj As Object)
    Private Function SaveJson(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        Dim act As Action(Of String, Object)
        act = Sub(json, obj)
                  Call obj.ModelSave(json, obj)
                  obj = obj.ModelLoad(json)
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
    Private Function RunRobot(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        Call rpa.ModelSave(rpa.SystemJsonFileName, rpa)
        Call rpa.MyProjectObject.SetData(trn, rpa)
        Dim i = rpa.MyProjectObject.Main()
        Call rpa.RootProjectObject.ModelSave(rpa.RootProjectJsonFileName, rpa.RootProjectObject)
        Call rpa.MyProjectObject.ModelSave(rpa.MyProjectJsonFileName, rpa.MyProjectObject)
        Return i
    End Function
    '---------------------------------------------------------------------------------------------'


    ' プロジェクトのアップデート
    '---------------------------------------------------------------------------------------------'
    Private Function UpdateProject(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        Dim i = -1
        Dim answer = vbNullString
        Dim pack As RpaPackage = Nothing

        If trn.Parameters.Count = 0 Then
            Console.WriteLine("パラメータが指定されていません: " & trn.CommandText)
            Return 1000
        End If

        If Not rpa.IsRootProjectExists Then
            Console.WriteLine("プロジェクト: '" & rpa.ProjectName & "' のルートプロジェクトが存在しません")
            Console.WriteLine("ダウンロード先が存在しません")
            Return 1000
        End If

        If rpa.RootProjectUpdatePackages.Count = 0 Then
            Console.WriteLine("'" & rpa.RootProjectUpdateDirectory & "' にパッケージリストが存在しません")
            Return 1000
        End If

        Console.WriteLine("パッケージを検索しています...")
        For Each p In rpa.RootProjectUpdatePackages
            If p.Name = trn.Parameters(0) Then
                pack = p
            End If
        Next
        If pack Is Nothing Then
            Console.WriteLine("指定のパッケージ '" & trn.Parameters(0) & "' はパッケージリストに登録されていません")
            Return 1000
        End If

        If pack.Latest Then
            answer = "y"
        Else
            Console.WriteLine("指定のパッケージ '" & pack.Name & "' は最新のパッケージではありません。")
            Console.WriteLine("既に前方互換性を持つパッケージをダウンロードしているか確認をしてください")
            Console.WriteLine(vbNullString)
            Console.WriteLine("ダウンロード済みパッケージ")
            Console.WriteLine("-----------------------------------------------------------")
            For Each p In rpa.MyProjectUpdatedPackages
                Console.WriteLine(p.Name)
            Next
            If rpa.MyProjectUpdatedPackages.Count = 0 Then
                Console.WriteLine("ダウンロード済みパッケージ　なし")
            End If
            Console.WriteLine("パッケージをダウンロードしますか？ (y/n)")
            Do
                Console.ReadLine()
            Loop Until answer = "y" Or "n"
        End If

        If answer = "n" Then
            Return 0
        End If

        Dim src = vbNullString
        Dim dsd = vbNullString
        Dim dst = vbNullString
        For Each ui In pack.UpdateInfos
            src = ui.SourceFile
            dsd = ui.DistinationSubDirectory
            If Not String.IsNullOrEmpty(dsd) Then
                dst = rpa.MyProjectDirectory & "\" & ui.DistinationSubDirectory & "\" & Path.GetFileName(src)
            Else
                dst = vbNullString
            End If

            If Not File.Exists(ui.SourceFile) Then
                Console.WriteLine("ソースファイル: " & src & " がありません")
                Console.WriteLine("詳しくは作成者に問い合わせてください... Author: " & ui.Author)
                Continue For
            End If

            If Not Directory.Exists(dsd) Then
                Console.WriteLine("アップデート先ディレクトリ: " & dsd & " がありません")
                Continue For
            End If

            File.Copy(src, dst, True)
        Next

        Return 0
    End Function
    '---------------------------------------------------------------------------------------------'

    ' プロジェクトのダウンロード
    '---------------------------------------------------------------------------------------------'
    Private Function DownloadProject(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        If trn.Parameters.Count > 0 Then
            Call SelectedDownload(trn, rpa)
        Else
            Call AllDownload(trn, rpa, rpa.RootProjectDirectory, rpa.MyProjectDirectory)
        End If

        Return 0
    End Function

    Private Sub SelectedDownload(ByRef trn As RpaTransaction, ByRef rpa As RpaProject)
        Dim src = vbNullString
        Dim dst = vbNullString
        Dim dstp = vbNullString
        For Each p In trn.Parameters
            src = rpa.RootProjectDirectory & "\" & p
            dst = rpa.MyProjectDirectory & "\" & p
            dstp = Path.GetDirectoryName(dst)
            If Not Directory.Exists(dstp) Then
                Console.WriteLine("ディレクトリが存在しません: " & dstp)
                Continue For
            End If
            If File.Exists(src) Then
                File.Copy(src, dst, True)
                Console.WriteLine("File.Copy... src: " & src)
                Console.WriteLine("          => dst: " & dst)
            ElseIf Directory.Exists(src) Then
                If Directory.Exists(dst) Then
                    Console.WriteLine("既にディレクトリが存在します: " & dst)
                Else
                    Directory.CreateDirectory(dst)
                    Console.WriteLine("Directory.CreateDirectory... src: " & src)
                    Console.WriteLine("                          => dst: " & dst)
                End If
            End If
        Next
    End Sub

    Private Sub AllDownload(ByRef trn As RpaTransaction, ByRef rpa As RpaProject,
                                  ByVal src As String, ByVal dst As String)
        If Not Directory.Exists(dst) Then
            Directory.CreateDirectory(dst)
            Console.WriteLine("Directory.CreateDirectory... src: " & src)
            Console.WriteLine("                          => dst: " & dst)
        End If

        Dim fsis As FileSystemInfo()
        Dim sdi = New DirectoryInfo(src)
        Dim ddi = New DirectoryInfo(dst)
        Dim src2 = vbNullString
        Dim dst2 = vbNullString
        fsis = sdi.GetFileSystemInfos
        For Each fsi In fsis
            src2 = fsi.FullName
            dst2 = ddi.FullName & "\" & fsi.Name
            If Not rpa.RootProjectIgnoreList.Contains(src2) Then
                If Not rpa.MyProjectIgnoreList.Contains(src2) Then
                    If ((fsi.Attributes And FileAttributes.Directory) = FileAttributes.Directory) Then
                        Call AllDownload(trn, rpa, src2, dst2)
                    Else
                        File.Copy(src2, dst2, True)
                        Console.WriteLine("File.Copy... src: " & src2)
                        Console.WriteLine("          => dst: " & dst2)
                    End If
                End If
            End If
        Next
    End Sub
    '---------------------------------------------------------------------------------------------'


    ' プロジェクトの切り替え
    '---------------------------------------------------------------------------------------------'
    Private Function SetProject(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        If trn.Parameters.Count = 0 Then
            Return 1000
        Else
            rpa = New RpaProject With {.ProjectName = trn.Parameters(0)}
            Return 0
        End If
    End Function
    '---------------------------------------------------------------------------------------------'


    ' プロジェクト終了
    '---------------------------------------------------------------------------------------------'
    Private Function RpaExit(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        trn.ExitFlag = True
        Return 0
    End Function
    '---------------------------------------------------------------------------------------------'

    Public Sub Main(ByRef trn As RpaTransaction, ByRef rpa As RpaProject)
        Dim fnc = ExecuteHandler(trn.MainCommand)
        If fnc IsNot Nothing Then
            trn.ReturnCode = fnc(trn, rpa)
        Else
            Console.WriteLine($"コマンド : '{trn.CommandText}' はありません")
        End If
    End Sub
End Class
