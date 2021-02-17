Public Class Rpa11 : Inherits Rpa00.RpaBase(Of Rpa11)

    Public Overrides Function SetupProjectObject(project As String) As Object
        Throw New NotImplementedException()
    End Function

    Private Const ENTER_KEY As String = "[ENTER]"

    Private Const CX_STAT_SYSTEM As Int32 = 1
    Private Const CX_STAT_INPERR As Int32 = 2
    Private Const CX_STAT_COMERR As Int32 = 4
    Private Const CX_STAT_PRINT As Int32 = 8

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

    Private Function Main(ByRef dat As Object) As Integer
        Dim dspemu As Object
        dspemu = CreateObject("Dspemu.Application")

        '-----------------------------------------------------------------------------------------'
        dspemu.WaitTime = 30
        dspemu.RectFlag = True
        dspemu.CopyMode = 1
        dspemu.SessionTo = "C:\Program Files (x86)\Fujitsu\WSMGR\PQ24DP002.emd"
        '-----------------------------------------------------------------------------------------'

        dspemu.Open
        dspemu.SendKeys("LOGON CONS")
        dspemu.SendKeys(ENTER_KEY)
        dspemu.Wait(5)

        dspemu.SendKeys("[PF1]")
        dspemu.Wait(1)

        dspemu.CopyMode = 0
        Dim i As Integer = dspemu.GetScreen()
        Dim txt As String = dspemu.ScreenData
        Console.WriteLine(txt)

        dspemu.CopyMode = 3
        Dim j As Integer = dspemu.GetScreen()
        Dim attr As String = dspemu.ScreenData
        Console.WriteLine(attr)

        dspemu.SendKeys(ENTER_KEY)

        dspemu.Wait(1)

        dspemu.SendKeys("LOGOFF")
        dspemu.SendKeys(ENTER_KEY)
        Dim k As Integer = dspemu.GetScreen()
        Console.WriteLine(dspemu.ScreenData)

        dspemu.Close
        dspemu.Quit
        dspemu = Nothing

        Return 0
    End Function

    Private Function Check(ByRef dat As Object) As Boolean
        Return True
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.CanExecuteHandler = AddressOf Check
    End Sub
End Class
