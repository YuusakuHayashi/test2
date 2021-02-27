Imports System.IO
Imports Newtonsoft.Json

Public Class Rpa11 : Inherits Rpa00.RpaBase(Of Rpa11)

    Public Overrides Function SetupProjectObject(project As String) As Object
        Throw New NotImplementedException()
    End Function

    Private Const ENTER_KEY As String = "[ENTER]"
    Private Const LOGON_KEY As String = "LOGON CONS"
    Private Const LOGOFF_KEY As String = "LOGOFF"
    Private Const PF1_KEY As String = "[PF1]"
    Private Const PF12_KEY As String = "[PF12]"

    Private Const CX_STAT_SYSTEM As Int32 = 1
    Private Const CX_STAT_INPERR As Int32 = 2
    Private Const CX_STAT_COMERR As Int32 = 4
    Private Const CX_STAT_PRINT As Int32 = 8

    Private Const GO_EXIT_SCREEN_TEXT = "--- *** END DISPLAY *** ---"
    Private Const GO_NEXT_SCREEN_TEXT = "--- *** CONTINUE *** ---"

    Private Enum OutputType
        BOTH = 0
        CONSOLE = 1
        FILE = 2
    End Enum

    Private Enum SystemConnection
        CX_SYS_NOTREADY = 0
        CX_SYS_READY = 1
        CX_SYS_BSC_CONNECT = 2
        CX_SYS_HDCL = 3
        CX_SYS_HDLC_CONNECT_NL = 4
        CX_SYS_HDLC_CONNECT = 5
        CX_SYS_DDX_READY = 6
        CX_SYS_DDX_CONNECT = 7
        CX_SYS_TEST = 8
        CX_SYS_OFFLINE = 9
    End Enum

    Private Enum InpStatus
        CX_INPCOM_UNLOCK = 0
        CX_INPCOM_LOCK = 1
        CX_INPCOM_LOCKKEY = 2
        CX_INPCOM_INVINPUT = 3
        CX_INPCOM_INVFUNC = 4
        CX_INPCOM_FIELDOVER = 5
        CX_INPCOM_INVPOS = 6
        CX_INPCOM_NUMONLY = 8
        CX_INPCOM_SHORTCOLUMN = 9
        CX_INPCOM_SHORTINPUT = 10
        CX_INPCOM_MISSINPUT = 11
        CX_INPCOM_MISSSPECIFY = 12
    End Enum

    Private Enum CommuError
        CX_COMMU_NORMAL = 0
        CX_COMMU_ERRPROG = 1
        CX_COMMU_ERRCOMMU = 2
        CX_COMMU_ERRNET = 3
        CX_COMMU_ERRCALL = 4
        CX_COMMU_PRTNODEDEVICE = 5
        CX_COMMU_PRTPRINTING = 6
        CX_COMMU_PRTUSING = 7
        CX_COMMU_PRTNOTPRINT = 8
        CX_COMMU_COMMUERROR = 9
    End Enum

    Private Enum PrinterStatus
        CX_PRNSTAT_NODEVICE = 0
        CX_PRNSTAT_READY = 1
        CX_PRNSTAT_READY_U = 2
        CX_PRNSTAT_PRINTING = 3
        CX_PRNSTAT_ERROR = 4
    End Enum

    Private _DspemuReturnCode As Int32
    Private Property DspemuReturnCode As Int32
        Get
            Return Me._DspemuReturnCode
        End Get
        Set(value As Int32)
            Me._DspemuReturnCode = value
        End Set
    End Property

    Private _OutputFileWriter As StreamWriter
    Private Property OutputFileWriter As StreamWriter
        Get
            If Me._OutputFileWriter Is Nothing Then
                Me._OutputFileWriter = New StreamWriter(Me.OutputFileName, False, Text.Encoding.GetEncoding("Shift-JIS"))
            End If
            Return Me._OutputFileWriter
        End Get
        Set(value As StreamWriter)
            Me._OutputFileWriter = value
        End Set
    End Property

    Private _LogFileWriter As StreamWriter
    Private Property LogFileWriter As StreamWriter
        Get
            If Me._OutputFileWriter Is Nothing Then
                Me._LogFileWriter = New StreamWriter(Me.LogFileName, False, Text.Encoding.GetEncoding("Shift-JIS"))
            End If
            Return Me._LogFileWriter
        End Get
        Set(value As StreamWriter)
            Me._LogFileWriter = value
        End Set
    End Property

    Private _TransactionLogs As List(Of String)
    Private Property TransactionLogs As List(Of String)
        Get
            If Me._TransactionLogs Is Nothing Then
                Me._TransactionLogs = New List(Of String)
            End If
            Return Me._TransactionLogs
        End Get
        Set(value As List(Of String))
            Me._TransactionLogs = value
        End Set
    End Property

    Private _ScreenDatas As List(Of String)
    Private ReadOnly Property ScreenDatas As List(Of String)
        Get
            If Me._ScreenDatas Is Nothing Then
                Me._ScreenDatas = New List(Of String)
            End If
            Return Me._ScreenDatas
        End Get
    End Property

    Private ReadOnly Property PageSeparater(ByVal page As Integer) As String
        Get
            Return $"                                                                        PAGE. {page}"
        End Get
    End Property

    Private ReadOnly Property LogSeparater() As String
        Get
            Return $"             < DSPEMU LOG >     LOGGED BY {Me.UserName}"
        End Get
    End Property

    ' ログファイル名
    Private ReadOnly Property LogFileName As String
        Get
            Dim yyyymmddhhmmss As String = Date.Now.ToString("yyyyMMddhhmmss")
            Return $"{Me._OutputLogDirectoryName}\log_{yyyymmddhhmmss}.txt"
        End Get
    End Property

    Private Enum DspemuLineType
        EJ_TITLE = 1
        EJ_WAITING_OUTPUT = 2
        EJ_JOB_HEADER = 3
        EJ_JOB_ROW2 = 4
        EJ_RTOW = 5
        EJ_END_DISPLAY = 10
        EJ_CONTINUE = 11
        EJ_OPERATOR_CALL = 99
    End Enum

    Private Enum DspemuLineColorStatus
        NORMAL = 0
        RED = 1
    End Enum

    Public Class DspemuLine
        Private _LineType As Integer
        Public Property LineType As Integer
            Get
                Return Me._LineType
            End Get
            Set(value As Integer)
                Me._LineType = value
            End Set
        End Property

        Private _Text As String
        Public Property Text As String
            Get
                Return Me._Text
            End Get
            Set(value As String)
                Me._Text = value
            End Set
        End Property

        Private _ColorStatus As Integer
        Public Property ColorStatus As Integer
            Get
                Return Me._ColorStatus
            End Get
            Set(value As Integer)
                Me._ColorStatus = value
            End Set
        End Property

        Private _JobStatus As String
        Public Property JobStatus As String
            Get
                Return Me._JobStatus
            End Get
            Set(value As String)
                Me._JobStatus = value
            End Set
        End Property
    End Class

    Private _DspemuPages As List(Of List(Of DspemuLine))
    Public Property DspemuPages As List(Of List(Of DspemuLine))
        Get
            If Me._DspemuPages Is Nothing Then
                Me._DspemuPages = New List(Of List(Of DspemuLine))
            End If
            Return Me._DspemuPages
        End Get
        Set(value As List(Of List(Of DspemuLine)))
            Me._DspemuPages = value
        End Set
    End Property

    ' Public 設定 --------------------------------------------------------------------------------'
    Private _UserName As String
    Public Property UserName As String
        Get
            Return Me._UserName
        End Get
        Set(value As String)
            Me._UserName = value
        End Set
    End Property

    Private _EMDFileName As String
    Public Property EMDFileName As String
        Get
            Return Me._EMDFileName
        End Get
        Set(value As String)
            Me._EMDFileName = value
        End Set
    End Property

    Private _OutputMode As Integer
    Public Property OutputMode As Integer
        Get
            Return Me._OutputMode
        End Get
        Set(value As Integer)
            Me._OutputMode = value
        End Set
    End Property

    Private _OutputFileName As String
    Public Property OutputFileName As String
        Get
            Return Me._OutputFileName
        End Get
        Set(value As String)
            Me._OutputFileName = value
        End Set
    End Property

    Private _SaveableLogsCount As Integer
    Public Property SaveableLogsCount As Integer
        Get
            Return Me._SaveableLogsCount
        End Get
        Set(value As Integer)
            Me._SaveableLogsCount = value
        End Set
    End Property

    Private _OutputLogDirectoryName As String
    Public Property OutputLogDirectoryName As String
        Get
            Return Me._OutputLogDirectoryName
        End Get
        Set(value As String)
            Me._OutputLogDirectoryName = value
        End Set
    End Property

    Private _DspemuWaitTime As Int32
    Public Property DspemuWaitTime As Int32
        Get
            Return Me._DspemuWaitTime
        End Get
        Set(value As Int32)
            Me._DspemuWaitTime = value
        End Set
    End Property

    Private _DspemuOpenMode As Int32
    Public Property DspemuOpenMode As Int32
        Get
            Return Me._DspemuOpenMode
        End Get
        Set(value As Int32)
            Me._DspemuOpenMode = value
        End Set
    End Property

    '---------------------------------------------------------------------------------------------'

    'Private Delegate Sub DspemuLineCheckDelegater(ByRef dline As DspemuLine)
    'Private _DspemuLineCheckHandler As List(Of DspemuLineCheckDelegater)
    'Private Property DspemuLineCheckHandler As List(Of DspemuLineCheckDelegater)
    '    Get
    '        If Me._DspemuLineCheckHandler Is Nothing Then
    '            Me._DspemuLineCheckHandler = New List(Of DspemuLineCheckDelegater)
    '        End If
    '        Return Me._DspemuLineCheckHandler
    '    End Get
    '    Set(value As List(Of DspemuLineCheckDelegater))
    '        Me._DspemuLineCheckHandler = value
    '    End Set
    'End Property

    Private Delegate Sub DspemuAutomationDelegater(ByRef dspemu As Object)

    ' Open
    '---------------------------------------------------------------------------------------------'
    Private _OpenHandler As List(Of DspemuAutomationDelegater)
    Private Property OpenHandler As List(Of DspemuAutomationDelegater)
        Get
            If Me._OpenHandler Is Nothing Then
                Me._OpenHandler = New List(Of DspemuAutomationDelegater)
            End If
            Return Me._OpenHandler
        End Get
        Set(value As List(Of DspemuAutomationDelegater))
            Me._OpenHandler = value
        End Set
    End Property

    Private Sub ConnectionEstablish(ByRef dspemu As Object)
        dspemu = GetObject(Me.EMDFileName)
        Me.DspemuReturnCode = dspemu.Open
    End Sub
    '---------------------------------------------------------------------------------------------'


    ' Main
    '---------------------------------------------------------------------------------------------'
    Private _MainHandler As List(Of DspemuAutomationDelegater)
    Private Property MainHandler As List(Of DspemuAutomationDelegater)
        Get
            If Me._MainHandler Is Nothing Then
                Me._MainHandler = New List(Of DspemuAutomationDelegater)
            End If
            Return Me._MainHandler
        End Get
        Set(value As List(Of DspemuAutomationDelegater))
            Me._MainHandler = value
        End Set
    End Property

    Private Sub SetConfiguration(ByRef dspemu As Object)
        dspemu.WaitTime = IIf(Me.DspemuWaitTime > 0, Me.DspemuWaitTime, dspemu.WaitTime)
        dspemu.RectFlag = True
        dspemu.OpenMode = Me.DspemuOpenMode
        dspemu.SessionTo = Me.EMDFileName
    End Sub

    Private Sub Logon(ByRef dspemu As Object)
        dspemu.WaitStatus(CX_STAT_INPERR, InpStatus.CX_INPCOM_UNLOCK, 1)
        Me.DspemuReturnCode = dspemu.SendKeys(LOGON_KEY)
        If Me.DspemuReturnCode > 0 Then
            Exit Sub
        End If

        Me.DspemuReturnCode = dspemu.SendKeys(ENTER_KEY)
        If Me.DspemuReturnCode > 0 Then
            Exit Sub
        End If
        dspemu.Wait(5)
    End Sub

    Private Sub GoExecutingJobMenu(ByRef dspemu As Object)
        dspemu.WaitStatus(CX_STAT_INPERR, InpStatus.CX_INPCOM_UNLOCK, 5)
        Me.DspemuReturnCode = dspemu.SendKeys(PF1_KEY)
        If Me.DspemuReturnCode > 0 Then
            Exit Sub
        End If
        dspemu.Wait(1)
    End Sub

    Private Sub GetScreen(ByRef dspemu As Object)
        Dim maxrow As Int32 = dspemu.DspemuRow
        Dim maxcol As Int32 = dspemu.DspemuColumn
        Do
            dspemu.WaitStatus(CX_STAT_INPERR, InpStatus.CX_INPCOM_UNLOCK, 1)
            dspemu.CopyMode = 0
            Dim row As Integer = 1
            Dim page As New List(Of DspemuLine)
            Dim repage As Boolean = False
            Do
                Me.DspemuReturnCode = dspemu.GetScreen(row, 1, row, maxcol)
                If Me.DspemuReturnCode > 0 Then
                    Exit Sub
                End If

                Dim dline As New DspemuLine With {.Text = dspemu.ScreenData}
                Call CheckDspemuLine(dline)
                page.Add(dline)

                If dline.LineType = DspemuLineType.EJ_END_DISPLAY Then
                    repage = False
                ElseIf dline.LineType = DspemuLineType.EJ_CONTINUE Then
                    repage = True
                Else
                End If

                row += 1
                If row > maxrow Then
                    Exit Do
                End If
            Loop Until False

            Me.DspemuPages.Add(page)

            If repage Then
                Me.DspemuReturnCode = dspemu.SendKeys(PF12_KEY)
                If Me.DspemuReturnCode > 0 Then
                    Exit Sub
                End If
                dspemu.Wait(1)
                Continue Do
            End If
            Exit Do
        Loop Until False
    End Sub

    Private Sub CheckDspemuLine(ByRef dline As DspemuLine)
        If Text.RegularExpressions.Regex.IsMatch(dline.Text, "OPERATOR CALL $") Then
            dline.LineType = DspemuLineType.EJ_OPERATOR_CALL
            dline.ColorStatus = DspemuLineColorStatus.RED
            Exit Sub
        End If
        If Text.RegularExpressions.Regex.IsMatch(dline.Text, "^ -+ *** JOB <EXECUTING> [0-9][0-9]\.[0-9][0-9] *** -^") Then
            dline.LineType = DspemuLineType.EJ_TITLE
            dline.ColorStatus = DspemuLineColorStatus.NORMAL
            Exit Sub
        End If
        If Text.RegularExpressions.Regex.IsMatch(dline.Text, "^ CLASS     ([A-Z],)*([A-Z])* WAITING OUTPUT") Then
            dline.LineType = DspemuLineType.EJ_WAITING_OUTPUT
            dline.ColorStatus = DspemuLineColorStatus.RED
            Exit Sub
        End If
        If Text.RegularExpressions.Regex.IsMatch(dline.Text, "^ [EHORD]\*[0-9]{3} [0-9]{2}*") Then
            dline.LineType = DspemuLineType.EJ_JOB_HEADER
            dline.ColorStatus = DspemuLineColorStatus.NORMAL
            dline.JobStatus = Strings.Mid(dline.Text, 2, 1)
            Exit Sub
        End If
        If Text.RegularExpressions.Regex.IsMatch(dline.Text, "^            T=[0-9][0-9]\.[0-9][0-9],RSIZE=\(") Then
            dline.LineType = DspemuLineType.EJ_JOB_ROW2
            dline.ColorStatus = DspemuLineColorStatus.NORMAL
            dline.JobStatus = Me.DspemuPages.Last.Last.JobStatus
            Exit Sub
        End If
        Dim pline As DspemuLine = Me.DspemuPages.Last.Last
        If pline.LineType = DspemuLineType.EJ_JOB_ROW2 And pline.JobStatus = "O" Then
            dline.LineType = DspemuLineType.EJ_RTOW
            dline.ColorStatus = DspemuLineColorStatus.RED
            dline.JobStatus = pline.JobStatus
            Exit Sub
        End If
        If Text.RegularExpressions.Regex.IsMatch(dline.Text, "^ -+ *** CONTINUE *** -+") Then
            dline.LineType = DspemuLineType.EJ_CONTINUE
            dline.ColorStatus = DspemuLineColorStatus.RED
            Exit Sub
        End If
        If Text.RegularExpressions.Regex.IsMatch(dline.Text, "^ -+ *** END DISPLAY *** -+") Then
            dline.LineType = DspemuLineType.EJ_END_DISPLAY
            dline.ColorStatus = DspemuLineColorStatus.NORMAL
            Exit Sub
        End If
    End Sub

    Private Sub LogOff(ByRef dspemu As Object)
        dspemu.WaitStatus(CX_STAT_INPERR, InpStatus.CX_INPCOM_UNLOCK, 1)
        Me.DspemuReturnCode = dspemu.SendKeys(ENTER_KEY)
        If Me.DspemuReturnCode > 0 Then
            Exit Sub
        End If
        dspemu.Wait(1)

        dspemu.WaitStatus(CX_STAT_INPERR, InpStatus.CX_INPCOM_UNLOCK, 1)
        Me.DspemuReturnCode = dspemu.SendKeys(LOGOFF_KEY)
        If Me.DspemuReturnCode > 0 Then
            Exit Sub
        End If
        Me.DspemuReturnCode = dspemu.SendKeys(ENTER_KEY)
        If Me.DspemuReturnCode > 0 Then
            Exit Sub
        End If
    End Sub

    'Private Sub Complete(ByRef dspemu As Object)
    '    Me.TransactionLogs.Add("DSPEMU操作は正常終了しました")
    'End Sub
    '---------------------------------------------------------------------------------------------'


    ' Close
    '---------------------------------------------------------------------------------------------'
    Private _CloseHandler As List(Of DspemuAutomationDelegater)
    Private Property CloseHandler As List(Of DspemuAutomationDelegater)
        Get
            If Me._CloseHandler Is Nothing Then
                Me._CloseHandler = New List(Of DspemuAutomationDelegater)
            End If
            Return Me._CloseHandler
        End Get
        Set(value As List(Of DspemuAutomationDelegater))
            Me._CloseHandler = value
        End Set
    End Property

    Private Sub Close(ByRef dspemu As Object)
        ' とりあえず、dspemu.Close() がおかしな復帰値を返したとしても
        ' dspemu.Quit() は実行してみる
        Me.DspemuReturnCode = dspemu.Close
        If Me.DspemuReturnCode = 0 Then
            Me.DspemuReturnCode = dspemu.Quit
        Else
            dspemu.Quit
        End If
    End Sub
    '---------------------------------------------------------------------------------------------'


    ' Output
    '---------------------------------------------------------------------------------------------'
    Private _OutputHandler As List(Of Action)
    Private Property OutputHandler As List(Of Action)
        Get
            If Me._OutputHandler Is Nothing Then
                Me._OutputHandler = New List(Of Action)
            End If
            Return Me._OutputHandler
        End Get
        Set(value As List(Of Action))
            Me._OutputHandler = value
        End Set
    End Property

    Private Sub WriteDspemuLine(ByVal txt As String)
        If Me.OutputMode = OutputType.BOTH Or Me.OutputMode = OutputType.FILE Then
            Me.OutputFileWriter.WriteLine(txt)
            Me.LogFileWriter.WriteLine(txt)
        ElseIf Me.OutputMode = OutputType.CONSOLE Then
            Console.WriteLine(txt)
        End If
    End Sub

    Private Sub PrintScreen()
        Try
            For Each page In Me.DspemuPages
                For Each line In page
                    Call WriteDspemuLine(line.Text)
                Next
                Call WriteDspemuLine(Me.PageSeparater(Me.ScreenDatas.IndexOf(Data) + 1))
                Call WriteDspemuLine($"{vbCrLf}{vbCrLf}{vbCrLf}{vbCrLf}")
            Next

            Call WriteDspemuLine(Me.LogSeparater)
            For Each log In Me.TransactionLogs
                Call WriteDspemuLine(log)
            Next
        Catch ex As Exception
        Finally
            Me.OutputFileWriter.Close()
            Me.OutputFileWriter.Dispose()
            Me.LogFileWriter.Close()
            Me.LogFileWriter.Dispose()
        End Try
    End Sub

    Private Sub DisposeLogs()
        Dim logs As List(Of String) = Directory.GetFiles(Me.OutputLogDirectoryName).ToList
        logs.Sort(
            Function(before, after)
                Return (before < after)
            End Function
        )
        Do
            If logs.Count <= Me.SaveableLogsCount Then
                Exit Do
            End If
            File.Delete(logs(0))
            logs.RemoveAt(0)
        Loop Until False
    End Sub
    '---------------------------------------------------------------------------------------------'

    ' DspemuAutomation プロジェクトから実行
    Public Function PrintExecutingJob() As Integer
        Me.OpenHandler.Add(AddressOf ConnectionEstablish)

        Me.MainHandler.Add(AddressOf SetConfiguration)
        Me.MainHandler.Add(AddressOf Logon)
        Me.MainHandler.Add(AddressOf GoExecutingJobMenu)
        Me.MainHandler.Add(AddressOf GetScreen)

        Me.CloseHandler.Add(AddressOf Close)

        Me.OutputHandler.Add(AddressOf PrintScreen)
        Me.OutputHandler.Add(AddressOf DisposeLogs)

        Return Main()
    End Function

    Private Overloads Function Main() As Integer
        Dim dspemu As Object
        Try
            For Each todo In Me.OpenHandler
                Call todo(dspemu)
                If Me.DspemuReturnCode > 0 Then
                    Dim i As Integer = dspemu.ErrorToMsg(Me.DspemuReturnCode)
                    Me.TransactionLogs.Add($"接続時のエラー ................... {Me.DspemuReturnCode} {dspemu.ErrorMsg}")
                    Exit For
                End If
                Me.TransactionLogs.Add($"接続が完了しました")
            Next
            Try
                If Me.DspemuReturnCode = 0 Then
                    Dim err As Boolean = False
                    For Each todo In Me.MainHandler
                        Call todo(dspemu)
                        If Me.DspemuReturnCode > 0 Then
                            Dim i As Integer = dspemu.ErrorToMsg(Me.DspemuReturnCode)
                            Me.TransactionLogs.Add($"主処理中のエラー ................. {Me.DspemuReturnCode} {dspemu.ErrorMsg}")
                            err = True
                            Exit For
                        End If
                    Next
                    If Not err Then
                        Me.TransactionLogs.Add($"主処理が完了しました")
                    End If
                End If
            Catch ex2 As Exception
                Me.TransactionLogs.Add($"主処理中の例外 ................... {ex2.HResult} {ex2.Message}")
            Finally
            End Try
        Catch ex As Exception
            Me.TransactionLogs.Add($"接続時の例外 ..................... {ex.HResult} {ex.Message}")
        Finally
            Try
                Dim err As Boolean = False
                For Each todo In Me.CloseHandler
                    Call todo(dspemu)
                    If Me.DspemuReturnCode > 0 Then
                        Dim i As Integer = dspemu.ErrorToMsg(Me.DspemuReturnCode)
                        Me.TransactionLogs.Add($"終了処理中のエラー ............... {Me.DspemuReturnCode} {dspemu.ErrorMsg}")
                        err = True
                    End If
                Next
                If Not err Then
                    Me.TransactionLogs.Add($"終了処理が完了しました")
                End If
            Catch ex3 As Exception
                Me.TransactionLogs.Add($"終了処理中の例外 ................. {ex3.HResult} {ex3.Message}")
            Finally
                dspemu = Nothing
                For Each todo In Me.OutputHandler
                    Call todo()
                Next

                Me.OpenHandler = Nothing
                Me.MainHandler = Nothing
                Me.CloseHandler = Nothing
                Me.OutputHandler = Nothing
            End Try
        End Try
        Return 0
    End Function

    ' RpaProject から実行
    Private Overloads Function Main(ByRef dat As Object) As Integer
        Me.OpenHandler.Add(AddressOf ConnectionEstablish)

        Me.MainHandler.Add(AddressOf SetConfiguration)
        Me.MainHandler.Add(AddressOf Logon)
        Me.MainHandler.Add(AddressOf GoExecutingJobMenu)
        Me.MainHandler.Add(AddressOf GetScreen)

        Me.CloseHandler.Add(AddressOf Close)

        Me.OutputHandler.Add(AddressOf PrintScreen)
        Me.OutputHandler.Add(AddressOf DisposeLogs)

        Return Main()
    End Function

    ' DspemuAutomation プロジェクトから実行
    Public Shared Function LoadFromFile(ByVal f As String) As Rpa11
        Dim rtn As Object
        Dim sr As StreamReader
        Try
            sr = New System.IO.StreamReader(f, System.Text.Encoding.GetEncoding("Shift-JIS"))
            Dim txt As String = sr.ReadToEnd()
            rtn = JsonConvert.DeserializeObject(Of Rpa11)(txt)
        Catch ex As Exception
            rtn = Nothing
            Console.WriteLine(ex.Message)
        Finally
            If sr IsNot Nothing Then
                sr.Close()
                sr.Dispose()
            End If
        End Try
        Return rtn
    End Function

    ' DspemuAutomation プロジェクトから実行
    Public Function CheckRun() As Boolean
        Return Check()
    End Function

    ' RpaProject プロジェクトから実行
    Private Overloads Function Check(ByRef dat As Object) As Boolean
        Return Check()
    End Function

    Private Overloads Function Check() As Boolean
        Dim b1 As Boolean = True
        If String.IsNullOrEmpty(Me.EMDFileName) Then
            Dim err As String = $"EMDFileName '{Me.EMDFileName}' が指定されていません"
            Console.WriteLine(err)
            b1 = False
        End If
        If Not File.Exists(Me.EMDFileName) Then
            Dim err As String = $"EMDFileName '{Me.EMDFileName}' は存在しません"
            Console.WriteLine(err)
            b1 = False
        End If
        If String.IsNullOrEmpty(Me.OutputFileName) Then
            Dim err As String = $"OutputFileName '{Me.OutputFileName}' が指定されていません"
            Console.WriteLine(err)
            b1 = False
        End If
        If String.IsNullOrEmpty(Me.OutputLogDirectoryName) Then
            Dim err As String = $"OutputLogDirectoryName '{Me.OutputLogDirectoryName}' が指定されていません"
            Console.WriteLine(err)
            b1 = False
        End If
        If Not Directory.Exists(Me.OutputLogDirectoryName) Then
            Dim err As String = $"OutputLogDirectoryName '{Me.OutputLogDirectoryName}' は存在しません"
            Console.WriteLine(err)
            b1 = False
        End If
        Return b1
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
    End Sub
End Class
