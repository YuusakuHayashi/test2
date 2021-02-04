Imports System.Runtime.InteropServices
Imports System.IO
Imports Rpa00
Imports Newtonsoft.Json

' Note : 
'     いつごろからか分からないが、マクロを呼び出ししようとすると、例外を出すようになった(0x8002005 (DISP_E_TYPEMISMATCH))
'     根本的な原因はわからないが、InvokeMacro(ByVal method As String, ByRef args() As Object) などは、
'     パラメタの要求として、オブジェクト型の配列を現状要求するようにしている。 
'     だが、配列の中身（各要素）がオブジェクト型ではない場合、この例外が発生することが分かった。
'     ただ、このクラス内にさらに内部クラス宣言している UpdateMacroCommand から InvokeMacro を呼び出す場合、
'     あからさまに文字列型の要素をセットしているが、例外は発生しない。
'     可能性として、内部クラスの場合、親クラスのメソッドの定義を知っているが、dllを外部から呼び出す場合は、
'     Object のため、メソッド定義が分からず、実行時例外が発生しているのか？・・・

Public Class RpaMacroUtility : Inherits RpaUtilityBase

    Friend Const MACROIMPORTER_BAS As String = "MacroImporter.bas"
    Friend Const MACROIMPORTER As String = "MacroImporter"
    Friend Const MACROIMPORTER2 As String = "MacroImporter2"
    Friend Const MAIN_FUNC As String = "Main"
    Friend Const SHOWMODULES_FUNC As String = "ShowModules"

    <JsonIgnore>
    Public Overrides ReadOnly Property CommandHandler(ByVal cmdtxt As String) As Object
        Get
            Dim cmd As Object
            Select Case cmdtxt
                Case "install" : cmd = New InstallMacroCommand(Me)
                Case "download" : cmd = New DownloadMacroFileCommand(Me)
                Case "update" : cmd = New UpdateMacroCommand(Me)
                Case "show" : cmd = New ShowMacrosCommand(Me)
                Case "push" : cmd = New PushMacrosCommand(Me)
                Case Else : cmd = Nothing
            End Select
            Return cmd
        End Get
    End Property

    <JsonIgnore>
    Public ReadOnly Property RootBasDirectory(dat As RpaDataWrapper) As String
        Get
            If Directory.Exists(dat.Project.RootDirectory) Then
                Dim [dir] As String = $"{dat.Project.RootDirectory}\bas"
                Return [dir]
            Else
                Return vbNullString
            End If
        End Get
    End Property

    <JsonIgnore>
    Public ReadOnly Property MacroFileName As String
        Get
            Dim f = $"{Me.MacroUtilityDirectory}\macro.xlsm"
            Return f
        End Get
    End Property

    Private _MacroUtilityDirectory As String
    Public Property MacroUtilityDirectory As String
        Get
            If String.IsNullOrEmpty(Me._MacroUtilityDirectory) Then
                Me._MacroUtilityDirectory = $"{RpaCui.SystemDirectory}\MacroUtility"
            End If
            If Not Directory.Exists(Me._MacroUtilityDirectory) Then
                Directory.CreateDirectory(Me._MacroUtilityDirectory)
            End If
            Return Me._MacroUtilityDirectory
        End Get
        Set(value As String)
            Me._MacroUtilityDirectory = value
        End Set
    End Property

    Private _ReleaseBasDirectory As String
    Public Property ReleaseBasDirectory As String
        Get
            If String.IsNullOrEmpty(Me._ReleaseBasDirectory) Then
                Me._ReleaseBasDirectory = vbNullString
            End If
            Return Me._ReleaseBasDirectory
        End Get
        Set(value As String)
            Me._ReleaseBasDirectory = value
        End Set
    End Property

    Private Function CallMacro(ByVal method As String, ByRef args() As Object) As Object
        Dim obj As Object
        Dim exapp, exbooks, exbook
        Dim macrofile As String = vbNullString

        Try
            exapp = CreateObject("Excel.Application")
            Try
                exbooks = exapp.Workbooks
                Try
                    exbook = exbooks.Open(Me.MacroFileName, 0)
                    macrofile = Path.GetFileName(Me.MacroFileName)
                    obj = exapp.Run($"{macrofile}!{method}", args)
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                    Console.WriteLine($"Workbookオブジェクトの取得、または、'{macrofile}!{method}' の実行に失敗しました")
                    obj = Nothing
                Finally
                    If exbook IsNot Nothing Then
                        exbook.Close()
                    End If
                    Marshal.ReleaseComObject(exbook)
                End Try
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Console.WriteLine($"Workbooksオブジェクトの取得に失敗しました")
                obj = Nothing
            Finally
                Marshal.ReleaseComObject(exbooks)
            End Try
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine($"Excel.Applicationオブジェクトの取得に失敗しました")
            obj = Nothing
        Finally
            If exapp IsNot Nothing Then
                exapp.Quit()
            End If
            Marshal.ReleaseComObject(exapp)
        End Try
        Return obj
    End Function

    Private Function ConvertArgsOfObjectType(ByRef args() As Object) As Object()
        Dim objs As New List(Of Object)
        For Each arg In args
            objs.Add(CType(arg, Object))
        Next
        Return (objs.ToArray())
    End Function

    Public Function InvokeMacro(ByVal method As String, ByRef args() As Object) As Integer
        Dim newargs() = ConvertArgsOfObjectType(args)
        Dim obj = Me.CallMacro(method, newargs)
        Return IIf((obj Is Nothing), 0, 1000)
    End Function

    Public Function InvokeMacroFunction(ByVal method As String, ByRef args() As Object) As Object
        Dim newargs() = ConvertArgsOfObjectType(args)
        Dim obj = Me.CallMacro(method, newargs)
        Return obj
    End Function


    ' Push Command
    '---------------------------------------------------------------------------------------------'
    Private Class PushMacrosCommand : Inherits RpaCommandBase
        Private Parent As RpaMacroUtility

        Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
            If String.IsNullOrEmpty(Parent.RootBasDirectory(dat)) Then
                Console.WriteLine($"'RootBasDirectory' は存在しません")
                Return False
            End If
            If Not Directory.Exists(Parent.RootBasDirectory(dat)) Then
                Console.WriteLine($"ディレクトリ '{Parent.RootBasDirectory(dat)}' は存在しません")
                Return False
            End If
            If String.IsNullOrEmpty(Parent.ReleaseBasDirectory) Then
                Console.WriteLine($"'ReleaseBasDirectory' は存在しません")
                Return False
            End If
            If Not Directory.Exists(Parent.ReleaseBasDirectory) Then
                Console.WriteLine($"ディレクトリ '{Parent.ReleaseBasDirectory}' は存在しません")
                Return False
            End If

            Return True
        End Function

        Private Function Main(ByRef dat As RpaDataWrapper) As Integer
            Dim srcdir As String = Parent.ReleaseBasDirectory
            Dim dstdir As String = Parent.RootBasDirectory(dat)
            For Each src In Directory.GetFiles(srcdir)
                Dim dst As String = $"{dstdir}\{Path.GetFileName(src)}"
                File.Copy(src, dst, True)
                Console.WriteLine($"Copy '{src}'")
                Console.WriteLine($" --> '{dst}'")
                Console.WriteLine()
            Next
            Return 0
        End Function

        Sub New(p As RpaMacroUtility)
            Me.Parent = p
            Me.ExecutableProjectArchitectures = {(New IntranetClientServerProject).GetType.Name}
            Me.ExecutableUser = {"RootDeveloper"}
            Me.ExecuteHandler = AddressOf Main
            Me.CanExecuteHandler = AddressOf Check
        End Sub
    End Class


    ' Show Command
    '---------------------------------------------------------------------------------------------'
    Private Class ShowMacrosCommand : Inherits RpaCommandBase
        Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
            If Not File.Exists(Me.Parent.MacroFileName) Then
                Console.WriteLine($"マクロファイル '{Me.Parent.MacroFileName}' が存在しません")
                Console.WriteLine($"開発者から入手してください")
                Console.WriteLine()
                Return False
            End If
            Return True
        End Function

        Private Parent As RpaMacroUtility
        Private Function Main(ByRef dat As RpaDataWrapper) As Integer
            Dim txt As String = (Parent.InvokeMacroFunction($"{MACROIMPORTER}.{SHOWMODULES_FUNC}", {""})).ToString.TrimEnd
            Console.WriteLine()
            Console.WriteLine($"インストール済マクロ一覧")
            Console.WriteLine($"{txt}")
            Console.WriteLine()
            Return 0
        End Function

        Sub New(p As RpaMacroUtility)
            Me.Parent = p
            Me.ExecuteHandler = AddressOf Main
            Me.CanExecuteHandler = AddressOf Check
        End Sub
    End Class


    ' Update Command
    '---------------------------------------------------------------------------------------------'
    Private Class UpdateMacroCommand : Inherits RpaCommandBase
        Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
            If Not File.Exists(Me.Parent.MacroFileName) Then
                Console.WriteLine($"マクロファイル '{Me.Parent.MacroFileName}' が存在しません")
                Console.WriteLine($"開発者から入手してください")
                Console.WriteLine()
                Return False
            End If

            If String.IsNullOrEmpty(dat.Project.RootDirectory) Then
                Console.WriteLine($"'RootDirectory' が設定されていません")
                Return False
            End If
            If Not Directory.Exists(dat.Project.RootDirectory) Then
                Console.WriteLine($"RootDirectory '{dat.Project.RootDirectory}' は存在しません")
                Return False
            End If
            If String.IsNullOrEmpty(dat.Project.RobotName) Then
                Console.WriteLine($"'RobotName' が設定されていません")
                Return False
            End If
            If Not Directory.Exists(dat.Project.RootRobotDirectory) Then
                Console.WriteLine($"RootRobotDirectory '{dat.Project.RootRobotDirectory}' は存在しません")
                Return False
            End If

            Dim basdir As String = Parent.RootBasDirectory(dat)
            If Not Directory.Exists(basdir) Then
                Console.WriteLine($"ディレクトリ '{basdir}' が存在しません")
                Console.WriteLine()
                Return False
            End If

            Return True
        End Function

        Private Parent As RpaMacroUtility

        Private Function Main(ByRef dat As RpaDataWrapper) As Integer
            Dim i As Integer = -1
            If dat.Transaction.Parameters.Count = 0 Then
                i = AllUpdate(dat)
            Else
                i = SelectedUpdate(dat)
            End If
            Return i
        End Function

        Private Function AllUpdate(ByRef dat As RpaDataWrapper) As Integer
            Dim utilpfx As String = $"{dat.Transaction.UtilityCommand}"
            dat.System.LateBindingCommands.Add($"{utilpfx} download")
            dat.System.LateBindingCommands.Add($"{utilpfx} install")
            Return 0
        End Function

        Private Function SelectedUpdate(ByRef dat As RpaDataWrapper) As Integer
            Dim utilpfx As String = $"{dat.Transaction.UtilityCommand}"
            dat.System.LateBindingCommands.Add($"{utilpfx} download {dat.Transaction.ParametersText}")
            dat.System.LateBindingCommands.Add($"{utilpfx} install {dat.Transaction.ParametersText}")
            Return 0
        End Function

        Sub New(p As RpaMacroUtility)
            Me.Parent = p
            Me.ExecuteHandler = AddressOf Main
            Me.CanExecuteHandler = AddressOf Check
            Me.ExecutableParameterCount = {0, 999}
        End Sub
    End Class


    ' Install Command
    '---------------------------------------------------------------------------------------------'
    Private Class InstallMacroCommand : Inherits RpaCommandBase
        Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
            If Not File.Exists(Me.Parent.MacroFileName) Then
                Console.WriteLine($"マクロファイル '{Me.Parent.MacroFileName}' が存在しません")
                Console.WriteLine($"開発者から入手してください")
                Console.WriteLine()
                Return False
            End If
            Return True
        End Function

        Private Parent As RpaMacroUtility

        Private Function Main(ByRef dat As RpaDataWrapper) As Integer
            Dim i As Integer = -1
            If dat.Transaction.Parameters.Count = 0 Then
                i = AllInstall(dat)
            Else
                i = SelectedInstall(dat)
            End If
            Return i
        End Function

        Private Function AllInstall(ByRef dat As RpaDataWrapper) As Integer
            Dim srcdir As String = Parent.MacroUtilityDirectory
            For Each src In Directory.GetFiles(srcdir)
                Dim ext As String = Path.GetExtension(src)
                Dim srcname As String = Path.GetFileName(src)
                If File.Exists(src) And (ext = ".bas" Or ext = ".cls") Then
                    Console.WriteLine($"マクロ '{srcname}' をインストールします")
                    If srcname = MACROIMPORTER_BAS Then
                        Call Parent.InvokeMacro($"{MACROIMPORTER2}.{MAIN_FUNC}", {src})
                    Else
                        Call Parent.InvokeMacro($"{MACROIMPORTER}.{MAIN_FUNC}", {src})
                    End If
                    Console.WriteLine()
                Else
                    Console.WriteLine($"マクロ '{srcname}' は存在しません")
                    Console.WriteLine()
                End If
            Next
            If Directory.GetFiles(srcdir).Count = 0 Then
                Console.WriteLine($"ディレクトリ '{srcdir}' に対象のマクロは存在しませんでした")
                Console.WriteLine()
            End If
            Return 0
        End Function

        Public Function SelectedInstall(ByRef dat As RpaDataWrapper) As Integer
            For Each para In dat.Transaction.Parameters
                Dim bas As String = $"{Parent.MacroUtilityDirectory}\{para}"
                Dim ext As String = Path.GetExtension(bas)
                If File.Exists(bas) And (ext = ".bas" Or ext = ".cls") Then
                    Console.WriteLine($"マクロ '{para}' をインストールします")
                    If para = MACROIMPORTER_BAS Then
                        Call Parent.InvokeMacro($"{MACROIMPORTER2}.{MAIN_FUNC}", {bas})
                    Else
                        Call Parent.InvokeMacro($"{MACROIMPORTER}.{MAIN_FUNC}", {bas})
                    End If
                    Console.WriteLine()
                Else
                    Console.WriteLine($"マクロ '{para}' は存在しません")
                    Console.WriteLine()
                End If
            Next
            Return 0
        End Function

        Sub New(p As RpaMacroUtility)
            Me.Parent = p
            Me.ExecuteHandler = AddressOf Main
            Me.CanExecuteHandler = AddressOf Check
            Me.ExecutableParameterCount = {0, 999}
        End Sub
    End Class


    ' Download
    '---------------------------------------------------------------------------------------------'
    Private Class DownloadMacroFileCommand : Inherits RpaCommandBase
        Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
            'If Not File.Exists(Me.Parent.MacroFileName) Then
            '    Console.WriteLine($"マクロファイル '{Me.Parent.MacroFileName}' が存在しません")
            '    Console.WriteLine($"開発者から入手してください")
            '    Console.WriteLine()
            '    Return False
            'End If

            If String.IsNullOrEmpty(dat.Project.RootDirectory) Then
                Console.WriteLine($"'RootDirectory' が設定されていません")
                Return False
            End If
            If Not Directory.Exists(dat.Project.RootDirectory) Then
                Console.WriteLine($"RootDirectory '{dat.Project.RootDirectory}' は存在しません")
                Return False
            End If
            If String.IsNullOrEmpty(dat.Project.RobotName) Then
                Console.WriteLine($"'RobotName' が設定されていません")
                Return False
            End If
            If Not Directory.Exists(dat.Project.RootRobotDirectory) Then
                Console.WriteLine($"RootRobotDirectory '{dat.Project.RootRobotDirectory}' は存在しません")
                Return False
            End If

            Dim basdir As String = Parent.RootBasDirectory(dat)
            If Not Directory.Exists(basdir) Then
                Console.WriteLine($"ディレクトリ '{basdir}' が存在しません")
                Console.WriteLine()
                Return False
            End If

            Return True
        End Function

        Private Parent As RpaMacroUtility

        Private Function Main(ByRef dat As RpaDataWrapper) As Integer
            Dim i As Integer = -1
            If dat.Transaction.Parameters.Count = 0 Then
                i = AllDownload(dat)
            Else
                i = SelectedDownload(dat)
            End If
            Return i
        End Function

        Private Function AllDownload(ByRef dat As RpaDataWrapper) As Integer
            Dim srcdir As String = Parent.RootBasDirectory(dat)
            Dim dstdir As String = Parent.MacroUtilityDirectory
            For Each src In Directory.GetFiles(srcdir)
                Dim ext As String = Path.GetExtension(src)
                Dim dst As String = $"{dstdir}\{Path.GetFileName(src)}"
                If File.Exists(src) And (ext = ".bas" Or ext = ".cls") Then
                    File.Copy(src, dst, True)
                    Console.WriteLine($"ファイルをコピー  '{src}'")
                    Console.WriteLine($"               => '{dst}'")
                    Console.WriteLine()
                Else
                    Console.WriteLine($"マクロ '{src}' は存在しません")
                    Console.WriteLine()
                End If
            Next
            If Directory.GetFiles(srcdir).Count = 0 Then
                Console.WriteLine($"ディレクトリ '{srcdir}' に対象のマクロは存在しませんでした")
                Console.WriteLine()
            End If
            Return 0
        End Function


        Private Function SelectedDownload(ByRef dat As RpaDataWrapper) As Integer
            Dim srcdir As String = Parent.RootBasDirectory(dat)
            Dim dstdir As String = Parent.MacroUtilityDirectory
            For Each para In dat.Transaction.Parameters
                Dim src As String = $"{srcdir}\{para}"
                Dim ext As String = Path.GetExtension(src)
                Dim dst As String = $"{dstdir}\{para}"
                If File.Exists(src) Then
                    File.Copy(src, dst, True)
                    Console.WriteLine($"ファイルをコピー  '{src}'")
                    Console.WriteLine($"               => '{dst}'")
                    Console.WriteLine()
                Else
                    Console.WriteLine($"マクロ '{src}' は存在しません")
                    Console.WriteLine()
                End If
            Next
            Return 0
        End Function

        Sub New(p As RpaMacroUtility)
            Me.Parent = p
            Me.ExecuteHandler = AddressOf Main
            Me.CanExecuteHandler = AddressOf Check
            Me.ExecutableParameterCount = {0, 999}
        End Sub
    End Class
End Class
