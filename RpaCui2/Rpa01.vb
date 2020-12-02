Imports System.IO

Public Class Rpa01 : Inherits RpaBase(Of Rpa01)
    Private Const EXCEL_APP As String = "Excel.Application"
    Private Const CSCRIPT As String = "cscript"

    Private _PasswordOfAttacheCase As String
    Public Property PasswordOfAttacheCase As String
        Get
            Return Me._PasswordOfAttacheCase
        End Get
        Set(value As String)
            Me._PasswordOfAttacheCase = value
        End Set
    End Property

    Private _OutlookSourceInboxName As String
    Public Property OutlookSourceInboxName As String
        Get
            Return Me._OutlookSourceInboxName
        End Get
        Set(value As String)
            Me._OutlookSourceInboxName = value
        End Set
    End Property

    Private _OutlookBackupInboxName As String
    Public Property OutlookBackupInboxName As String
        Get
            Return Me._OutlookBackupInboxName
        End Get
        Set(value As String)
            Me._OutlookBackupInboxName = value
        End Set
    End Property

    Private _MasterCsvFileName As String
    Public Property MasterCsvFileName As String
        Get
            If String.IsNullOrEmpty(Me._MasterCsvFileName) Then
                Me._MasterCsvFileName = "master.csv"
            End If
            Return Me._MasterCsvFileName
        End Get
        Set(value As String)
            Me._MasterCsvFileName = value
        End Set
    End Property

    Private _AttacheCase As String
    Public Property AttacheCase As String
        Get
            Return Me._AttacheCase
        End Get
        Set(value As String)
            Me._AttacheCase = value
        End Set
    End Property

    Private _RestartCode As String
    Public Property RestartCode As String
        Get
            Return Me._RestartCode
        End Get
        Set(value As String)
            Me._RestartCode = value
        End Set
    End Property

    Public Overrides Function SetupProjectObject(project As String) As Object
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Main() As Integer
        Dim master = Rpa.MyProjectDirectory & "\" & Me.MasterCsvFileName
        Console.WriteLine("以下のファイルを用意したら、[Enter]キーをクリックしてください")
        Console.WriteLine("ファイル名 : " & Me.MasterCsvFileName)
        Console.ReadLine()
        If Not File.Exists(master) Then
            Console.WriteLine("エラー：ファイル '${master}' が存在しません")
            Return 1000
        End If

        ' ワークディレクトリの作成
        If Not Directory.Exists(Rpa.MyProjectWorkDirectory) Then
            Directory.CreateDirectory(Rpa.MyProjectWorkDirectory)
        End If

        Dim wmaster = Rpa.MyProjectWorkDirectory & "\" & Me.MasterCsvFileName
        File.Copy(master, wmaster, True)

        Dim wtemp = Rpa.MyProjectWorkDirectory & "\tmp.csv"

        Console.WriteLine("添付ファイルを検索しています...")
        Call _GetAttachmentFile()

        Dim bname = vbNullString
        Dim dfile = vbNullString
        For Each f In Directory.GetFiles(Rpa.MyProjectWorkDirectory)
            If Path.GetExtension(f) Then
                dfile = f
            End If
        Next
        If String.IsNullOrEmpty(dfile) Then
            Console.WriteLine("エラー：添付ファイルが見つかりません")
            Return 1000
        End If
        bname = Path.GetFileNameWithoutExtension(dfile)

        Console.WriteLine("添付ファイルを解凍します...")
        If File.Exists(Me.AttacheCase) Then
            Console.WriteLine("エラー：アタッシュケース '${Me.AttacheCase}' がありません")
            Return 1000
        End If

        Call _RunCmd("/c " _
                   & "/p=${Me.AttacheCase} " _
                   & "/de=1 /ow=0 /opf=0 /exit=1")

        Dim infile = Rpa.MyProjectWorkDirectory & "\" & bname & ".xls"
        Dim wkfile = Rpa.MyProjectWorkDirectory & "\input.xls"
        If Not File.Exists(infile) Then
            Console.WriteLine("エラー：解凍後ファイル '${infile}' がありません")
            Return 1000
        End If
        File.Copy(infile, wkfile, True)

        Call _RunCmd("/c " _
                   & "${CSCRIPT} ${Rpa.MyProjectScriptDirectory}\${")
    End Function

    Private Sub _RunCmd(ByVal arg As String)
        Dim proc = New System.Diagnostics.Process()
        proc.StartInfo.FileName = System.Environment.GetEnvironmentVariable("ComSpec")
        proc.StartInfo.UseShellExecute = False
        proc.StartInfo.RedirectStandardOutput = True
        proc.StartInfo.RedirectStandardInput = False
        proc.StartInfo.CreateNoWindow = True
        proc.StartInfo.Arguments = arg
        proc.Start()
        proc.WaitForExit()
        proc.Close()
    End Sub

    Private Sub _GetAttachmentFile()
        Const MAPI = "MAPI"
        Const HIZUKE = "日付"
        Dim ol = CreateObject("Outlook.Application")
        Dim ns = ol.GetNamespace(MAPI)
        Dim inbox = ns.GetDefaultFolder(6)
        Dim sfolder = inbox.Folders.Item(Me.OutlookSourceInboxName)
        Dim bfolder = inbox.Folders.Item(Me.OutlookBackupInboxName)
        Dim afile As String

        sfolder.Items.Sort(HIZUKE, True)
        For Each [item] In sfolder.Items
            If [item].Attachments.Count = 0 Then
                Continue For
            End If
            For Each attachment In [item].Attachments
                afile = Rpa.MyProjectWorkDirectory & "\" & attachment
                attachment.SaveAsFile(afile)
            Next
        Next

        Do Until sfolder.Items.Count = 0
            sfolder.Items(0).Move(bfolder)
        Loop
    End Sub
End Class
