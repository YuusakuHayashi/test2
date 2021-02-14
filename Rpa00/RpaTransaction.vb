Imports System.IO
Public Class RpaTransaction
    'Public ReadOnly Property CommandTextLogFileName As String
    '    Get
    '        Dim [fil] As String = $"{RpaCui.SystemDirectory}\command.log"
    '        Dim jh As New Rpa00.JsonHandler(Of List(Of CommandLogData))
    '        If Not File.Exists([fil]) Then
    '            Call jh.Save(Of List(Of CommandLogData))([fil], (New List(Of CommandLogData)))
    '        End If
    '        Return [fil]
    '    End Get
    'End Property

    'Public ReadOnly Property SystemLogsDirectory As String
    '    Get
    '        Dim [dir] As String = $"{RpaCui.SystemDirectory}\logs"
    '        If Not Directory.Exists([dir]) Then
    '            Directory.CreateDirectory([dir])
    '        End If
    '        Return [dir]
    '    End Get
    'End Property

    'Private _SystemMonthlyLogDirectories As String()
    'Public ReadOnly Property SystemMonthlyLogDirectories As String()
    '    Get
    '        Dim dirs(12) As String
    '        If Me._SystemMonthlyLogDirectories Is Nothing Then
    '            dirs(0) = vbNullString
    '            For i As Integer = 1 To 12
    '                dirs(i) = $"{Me.SystemLogsDirectory}\{String.Format("{0:00}", i)}"
    '                If Not Directory.Exists(dirs(i)) Then
    '                    Directory.CreateDirectory(dirs(i))
    '                End If
    '            Next
    '            Me._SystemMonthlyLogDirectories = dirs
    '        End If
    '        Return Me._SystemMonthlyLogDirectories
    '    End Get
    'End Property

    'Private _SystemMonthlyLogFiles As String()
    'Public ReadOnly Property SystemMonthlyLogFiles As String()
    '    Get
    '        Dim fils(12) As String
    '        If Me._SystemMonthlyLogFiles Is Nothing Then
    '            fils(0) = vbNullString
    '            For i As Integer = 1 To 12
    '                fils(i) = $"{Me.SystemMonthlyLogDirectories(i)}\commands.log"
    '            Next
    '            Me._SystemMonthlyLogFiles = fils
    '        End If
    '        Return Me._SystemMonthlyLogFiles
    '    End Get
    'End Property

    'Private _ProjectArchitecture As Integer
    'Public Property ProjectArchitecture As Integer
    '    Get
    '        Return Me._ProjectArchitecture
    '    End Get
    '    Set(value As Integer)
    '        Me._ProjectArchitecture = value
    '    End Set
    'End Property

    'Private _CommandText As String
    'Public Property CommandText As String
    '    Get
    '        Return Me._CommandText
    '    End Get
    '    Set(value As String)
    '        Me._CommandText = value
    '    End Set
    'End Property

    'Private _IsAutoCommand As Boolean
    'Public Property IsAutoCommand As Boolean
    '    Get
    '        Return Me._IsAutoCommand
    '    End Get
    '    Set(value As Boolean)
    '        Me._IsAutoCommand = value
    '    End Set
    'End Property

    'Private _UtilityCommand As String
    'Public Property UtilityCommand As String
    '    Get
    '        Return Me._UtilityCommand
    '    End Get
    '    Set(value As String)
    '        Me._UtilityCommand = value
    '    End Set
    'End Property

    'Private _TrueCommand As String
    'Public Property TrueCommand As String
    '    Get
    '        Return Me._TrueCommand
    '    End Get
    '    Set(value As String)
    '        Me._TrueCommand = value
    '    End Set
    'End Property

    'Private _MainCommand As String
    'Public Property MainCommand As String
    '    Get
    '        Return Me._MainCommand
    '    End Get
    '    Set(value As String)
    '        Me._MainCommand = value
    '    End Set
    'End Property

    'Private _Modes As List(Of String)
    'Public Property Modes As List(Of String)
    '    Get
    '        If Me._Modes Is Nothing Then
    '            Me._Modes = New List(Of String)
    '        End If
    '        Return Me._Modes
    '    End Get
    '    Set(value As List(Of String))
    '        Me._Modes = value
    '    End Set
    'End Property

    'Private _Parameters As List(Of String)
    'Public Property Parameters As List(Of String)
    '    Get
    '        If Me._Parameters Is Nothing Then
    '            Me._Parameters = New List(Of String)
    '        End If
    '        Return Me._Parameters
    '    End Get
    '    Set(value As List(Of String))
    '        Me._Parameters = value
    '    End Set
    'End Property

    'Private _ParametersText As String
    'Public Property ParametersText As String
    '    Get
    '        Return Me._ParametersText
    '    End Get
    '    Set(value As String)
    '        Me._ParametersText = value
    '    End Set
    'End Property

    ''Private _ExitFlag As Boolean
    ''Public Property ExitFlag As Boolean
    ''    Get
    ''        Return Me._ExitFlag
    ''    End Get
    ''    Set(value As Boolean)
    ''        Me._ExitFlag = value
    ''    End Set
    ''End Property

    'Private _CommandLogs As List(Of CommandLogData)
    'Public Property CommandLogs As List(Of CommandLogData)
    '    Get
    '        If Me._CommandLogs Is Nothing Then
    '            Me._CommandLogs = New List(Of CommandLogData)
    '        End If
    '        Return Me._CommandLogs
    '    End Get
    '    Set(value As List(Of CommandLogData))
    '        Me._CommandLogs = value
    '    End Set
    'End Property

    'Public Class CommandLogData
    '    Private _UserName As String
    '    Public Property UserName As String
    '        Get
    '            Return Me._UserName
    '        End Get
    '        Set(value As String)
    '            Me._UserName = value
    '        End Set
    '    End Property

    '    Private _ProjectArchTypeName As String
    '    Public Property ProjectArchTypeName As String
    '        Get
    '            Return Me._ProjectArchTypeName
    '        End Get
    '        Set(value As String)
    '            Me._ProjectArchTypeName = value
    '        End Set
    '    End Property

    '    Private _ProjectName As String
    '    Public Property ProjectName As String
    '        Get
    '            Return Me._ProjectName
    '        End Get
    '        Set(value As String)
    '            Me._ProjectName = value
    '        End Set
    '    End Property

    '    Private _RobotName As String
    '    Public Property RobotName As String
    '        Get
    '            Return Me._RobotName
    '        End Get
    '        Set(value As String)
    '            Me._RobotName = value
    '        End Set
    '    End Property

    '    Private _RunDate As String
    '    Public Property RunDate As String
    '        Get
    '            Return Me._RunDate
    '        End Get
    '        Set(value As String)
    '            Me._RunDate = value
    '        End Set
    '    End Property

    '    Private _CommandText As String
    '    Public Property CommandText As String
    '        Get
    '            Return Me._CommandText
    '        End Get
    '        Set(value As String)
    '            Me._CommandText = value
    '        End Set
    '    End Property

    '    Private _UtilityCommand As String
    '    Public Property UtilityCommand As String
    '        Get
    '            Return Me._UtilityCommand
    '        End Get
    '        Set(value As String)
    '            Me._UtilityCommand = value
    '        End Set
    '    End Property

    '    Private _Result As String
    '    Public Property Result As String
    '        Get
    '            Return Me._Result
    '        End Get
    '        Set(value As String)
    '            Me._Result = value
    '        End Set
    '    End Property

    '    Private _ResultString As String
    '    Public Property ResultString As String
    '        Get
    '            Return Me._ResultString
    '        End Get
    '        Set(value As String)
    '            Me._ResultString = value
    '        End Set
    '    End Property

    '    Private _AutoCommandFlag As Boolean
    '    Public Property AutoCommandFlag As Boolean
    '        Get
    '            Return Me._AutoCommandFlag
    '        End Get
    '        Set(value As Boolean)
    '            Me._AutoCommandFlag = value
    '        End Set
    '    End Property

    '    Private _ExecuteTime As TimeSpan
    '    Public Property ExecuteTime As TimeSpan
    '        Get
    '            Return Me._ExecuteTime
    '        End Get
    '        Set(value As TimeSpan)
    '            Me._ExecuteTime = value
    '        End Set
    '    End Property
    'End Class

    'Public Sub SaveCommandLogs()
    '    Dim sw As New StreamWriter(
    '        Me.CommandTextLogFileName,
    '        True,
    '        Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING)
    '    )

    '    For Each [log] In Me.CommandLogs
    '        Dim line As String = vbNullString
    '        line &= $"{[log].RunDate},"
    '        line &= $"{[log].UserName},"
    '        line &= $"{[log].ProjectArchTypeName},"
    '        line &= $"{[log].ProjectName},"
    '        line &= $"{[log].RobotName},"
    '        line &= $"{[log].CommandText},"
    '        line &= $"{[log].UtilityCommand},"
    '        line &= $"{[log].Result},"
    '        line &= $"{[log].ResultString},"
    '        line &= $"{IIf([log].AutoCommandFlag, "1", "0")},"
    '        line &= $"{[log].ExecuteTime.ToString}"
    '        sw.WriteLine(line)
    '    Next

    '    sw.Close()
    '    sw.Dispose()

    '    Me.CommandLogs = New List(Of CommandLogData)
    'End Sub

    Public Function ShowRpaIndicator(ByRef dat As RpaDataWrapper) As String
        ' ガイド
        If dat.Project IsNot Nothing Then
            Console.Write($"{dat.Project.GuideString}")
        Else
            Console.Write("NoRpa>")
        End If

        Return ShowReader(dat)
    End Function

    Private Function ShowReader(ByRef dat As RpaDataWrapper) As String
        ' 通常
        If Not dat.Initializer.ReleaseVersion = "Experimental" Then
            Return Console.ReadLine()
        End If

        ' 開発中
        Dim txt As String = vbNullString
        Dim cmptxt As String = vbNullString
        Dim cnt As Integer = 0
        Do
            Dim cki As ConsoleKeyInfo = Console.ReadKey(True)

            Select Case cki.Key
                ' 1-9, A-Z
                Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57,
                     65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90,
                     96, 97, 98, 99, 100, 101, 102, 103, 104, 105
                    Console.Write(cki.KeyChar)
                    txt &= cki.KeyChar
                ' Space
                Case ConsoleKey.Spacebar
                    Console.Write(" "c)
                    txt &= cki.KeyChar
                ' BackSpace
                Case ConsoleKey.Backspace
                    If Not String.IsNullOrEmpty(txt) Then
                        'Console.Write(cki.KeyChar)
                        Console.Write(vbBack)
                        Console.Write(" "c)
                        Console.Write(vbBack)
                        txt = txt.Trim(txt.Last)
                    End If
                ' Tab
                Case ConsoleKey.Tab
                    cmptxt = ComplementCommandFromSnippet(dat, txt)
                    txt &= cmptxt
                'Case ConsoleKey.DownArrow, ConsoleKey.UpArrow
                '    Console.Write(cki.KeyChar)
                ' Enter
                Case ConsoleKey.Enter
                    txt &= cki.KeyChar
                Case Else
                    ' Nothing To Do
            End Select

            If String.IsNullOrEmpty(txt) Then
                Continue Do
            End If

            If txt.Last = vbCr Then
                Console.WriteLine()
                txt = txt.Trim(vbCr)
                Exit Do
            End If
        Loop Until False

        Return txt
    End Function

    ' Tab押下によるスニペット機能
    Private Function ComplementCommandFromSnippet(ByRef dat As RpaDataWrapper, ByVal txt As String) As String
        Dim baselen As Integer = txt.Length
        Dim tabcnt As Integer = 0
        Dim rtntxt As String = vbNullString

        If baselen > 0 Then
            Do
                Dim cmptxt As String = vbNullString    ' 補完する文字列（例：入力 'ex' => cmptxt = 'ix')
                Dim cmplen As Integer = 0              ' 補完する文字列の長さ
                Dim hitcnt As Integer = 0              ' 辞書からヒットした数
                For Each cmdpair In dat.System.CommandDictionary
                    If Strings.Left(cmdpair.Key, baselen) = txt Then
                        hitcnt += 1
                        If tabcnt >= hitcnt Then
                            Continue For
                        End If

                        cmptxt = cmdpair.Key.Substring(baselen)
                        cmplen = cmptxt.Length()
                        Console.Write(cmptxt)
                        Exit For
                    End If
                Next

                If String.IsNullOrEmpty(cmptxt) Then
                    Exit Do
                End If

                ' 連続タブ入力
                Dim cki As ConsoleKeyInfo = Console.ReadKey(True)
                Select Case cki.Key
                    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57,
                         65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90,
                         96, 97, 98, 99, 100, 101, 102, 103, 104, 105
                        ' 1-9, A-Z
                        Console.Write(cki.KeyChar)
                        rtntxt = cmptxt & cki.KeyChar
                        Exit Do
                    Case ConsoleKey.Backspace
                        'Console.Write(cki.KeyChar)
                        Console.Write(vbBack)
                        Console.Write(" "c)
                        Console.Write(vbBack)
                        rtntxt = cmptxt.TrimEnd(cmptxt.Last)
                        Exit Do
                    Case ConsoleKey.Spacebar
                        Console.Write(cki.KeyChar)
                        rtntxt = cmptxt & cki.KeyChar
                        Exit Do
                    Case ConsoleKey.Tab
                        Console.Write(Strings.StrDup(cmplen, vbBack))
                        Console.Write(Strings.StrDup(cmplen, " "c))
                        Console.Write(Strings.StrDup(cmplen, vbBack))
                        tabcnt += 1
                        Continue Do
                    Case ConsoleKey.Escape
                        Console.Write(Strings.StrDup(cmplen, vbBack))
                        Console.Write(Strings.StrDup(cmplen, " "c))
                        Console.Write(Strings.StrDup(cmplen, vbBack))
                        rtntxt = vbNullString
                        Exit Do
                    Case ConsoleKey.Enter
                        rtntxt = cmptxt & cki.KeyChar
                        Exit Do
                    Case Else
                        Exit Do
                End Select
            Loop Until False
        End If

        Return rtntxt
    End Function
End Class
