Imports System.IO
Public Class RpaTransaction
    Public ReadOnly Property CommandTextLogFileName As String
        Get
            Dim [fil] As String = $"{RpaCui.SystemDirectory}\command.log"
            Dim jh As New RpaCui.JsonHandler(Of List(Of CommandLogData))
            If Not File.Exists([fil]) Then
                Call jh.Save(Of List(Of CommandLogData))([fil], (New List(Of CommandLogData)))
            End If
            Return [fil]
        End Get
    End Property

    Private _ReturnCode As Integer
    Public Property ReturnCode As Integer
        Get
            Return Me._ReturnCode
        End Get
        Set(value As Integer)
            Me._ReturnCode = value
        End Set
    End Property

    Private _ProjectArchitecture As Integer
    Public Property ProjectArchitecture As Integer
        Get
            Return Me._ProjectArchitecture
        End Get
        Set(value As Integer)
            Me._ProjectArchitecture = value
        End Set
    End Property

    Private _CommandText As String
    Public Property CommandText As String
        Get
            Return Me._CommandText
        End Get
        Set(value As String)
            Me._CommandText = value
        End Set
    End Property

    Private _MainCommand As String
    Public Property MainCommand As String
        Get
            Return Me._MainCommand
        End Get
        Set(value As String)
            Me._MainCommand = value
        End Set
    End Property

    Private _Modes As List(Of String)
    Public Property Modes As List(Of String)
        Get
            If Me._Modes Is Nothing Then
                Me._Modes = New List(Of String)
            End If
            Return Me._Modes
        End Get
        Set(value As List(Of String))
            Me._Modes = value
        End Set
    End Property

    Private _Parameters As List(Of String)
    Public Property Parameters As List(Of String)
        Get
            If Me._Parameters Is Nothing Then
                Me._Parameters = New List(Of String)
            End If
            Return Me._Parameters
        End Get
        Set(value As List(Of String))
            Me._Parameters = value
        End Set
    End Property

    Private _ExitFlag As Boolean
    Public Property ExitFlag As Boolean
        Get
            Return Me._ExitFlag
        End Get
        Set(value As Boolean)
            Me._ExitFlag = value
        End Set
    End Property

    Private _CommandLogs As List(Of CommandLogData)
    Public Property CommandLogs As List(Of CommandLogData)
        Get
            If Me._CommandLogs Is Nothing Then
                Me._CommandLogs = New List(Of CommandLogData)
            End If
            Return Me._CommandLogs
        End Get
        Set(value As List(Of CommandLogData))
            Me._CommandLogs = value
        End Set
    End Property

    Public Class CommandLogData
        Private _RunDate As String
        Public Property RunDate As String
            Get
                If String.IsNullOrEmpty(Me._RunDate) Then
                    Me._RunDate = (DateTime.Now).ToString
                End If
                Return Me._RunDate
            End Get
            Set(value As String)
                Me._RunDate = value
            End Set
        End Property

        Private _CommandText As String
        Public Property CommandText As String
            Get
                Return Me._CommandText
            End Get
            Set(value As String)
                Me._CommandText = value
            End Set
        End Property
    End Class

    Public Sub SaveCommandLogs()
        Dim jh As New RpaCui.JsonHandler(Of List(Of CommandLogData))
        Dim [olds] As List(Of CommandLogData) = jh.Load(Of List(Of CommandLogData))(Me.CommandTextLogFileName)
        For Each [new] In Me.CommandLogs
            olds.Add([new])
        Next
        Call jh.Save(Of List(Of CommandLogData))(Me.CommandTextLogFileName, [olds])
        Me.CommandLogs = New List(Of CommandLogData)
    End Sub

    Public Sub CreateCommand()
        Dim texts() As String
        Me.MainCommand = vbNullString
        Me.Parameters = New List(Of String)

        texts = Me.CommandText.Split(" ")
        Me.MainCommand = texts(0)

        Dim i As Integer = 0
        For Each p In texts
            If i <> 0 Then
                Me.Parameters.Add(p)
            End If
            i += 1
        Next
    End Sub

    Public Function ShowRpaIndicator(ByRef dat As RpaDataWrapper) As String
        ' ガイド
        If Me.Modes.Count = 0 Then
            If dat.Project IsNot Nothing Then
                Console.Write($"{dat.Project.ProjectName}\{dat.Project.RobotAlias}>")
            Else
                Console.Write("NoRpa>")
            End If
        Else
            For Each mode In Me.Modes
                If Me.Modes.Last = mode Then
                    Console.Write($"{mode}>")
                Else
                    Console.Write($"{mode}\")
                End If
            Next
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
