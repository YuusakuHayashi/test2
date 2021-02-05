Imports System.IO

Public Class HelpCommand : Inherits RpaCommandBase

    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        Return True
    End Function

    Private ReadOnly Property Navigater As String
        Get
            Dim pstr As String = $"p. {Me.Page.ToString}"
            Dim pstrlen As Integer = pstr.Length
            Dim bstr As String = IIf(Me.Page = Me.BeginPage, "    ", "  <<")
            Dim bstrlen As Integer = bstr.Length
            Dim fstr As String = IIf(Me.Page = Me.LastPage, "END ", ">> ")
            Dim fstrlen As Integer = fstr.Length

            Dim spclen As Integer = Console.WindowWidth - (pstrlen + bstrlen + fstrlen)
            Dim lspclen As Integer = 0
            Dim rspclen As Integer = 0

            If (spclen Mod 2 <> 0) Then
                rspclen = (spclen / 2)
                lspclen = (rspclen + 1)
            Else
                rspclen = spclen / 2
                lspclen = rspclen
            End If

            Dim navstr As String = $"{bstr}{Strings.StrDup(lspclen, " "c)}{pstr}{Strings.StrDup(rspclen, " "c)}{fstr}"
            Return navstr
        End Get
    End Property

    Private _ForwardChars As List(Of ConsoleKey)
    Private ReadOnly Property ForwardChars As List(Of ConsoleKey)
        Get
            If Me._ForwardChars Is Nothing Then
                Me._ForwardChars = New List(Of ConsoleKey)
                Me._ForwardChars.Add(ConsoleKey.RightArrow)
                Me._ForwardChars.Add(ConsoleKey.L)
                Me._ForwardChars.Add(ConsoleKey.PageDown)
                Me._ForwardChars.Add(ConsoleKey.Enter)
                Me._ForwardChars.Add(ConsoleKey.Add)
            End If
            Return Me._ForwardChars
        End Get
    End Property

    Private _BackChars As List(Of ConsoleKey)
    Private ReadOnly Property BackChars As List(Of ConsoleKey)
        Get
            If Me._BackChars Is Nothing Then
                Me._BackChars = New List(Of ConsoleKey)
                Me._BackChars.Add(ConsoleKey.LeftArrow)
                Me._BackChars.Add(ConsoleKey.H)
                Me._BackChars.Add(ConsoleKey.PageUp)
                Me._BackChars.Add(ConsoleKey.Backspace)
                Me._ForwardChars.Add(ConsoleKey.Subtract)
            End If
            Return Me._BackChars
        End Get
    End Property

    Private _Reader As StreamReader
    Private Property Reader As StreamReader
        Get
            Return Me._Reader
        End Get
        Set(value As StreamReader)
            Me._Reader = value
        End Set
    End Property

    Private _HelpFile As String
    Private Property HelpFile As String
        Get
            Return Me._HelpFile
        End Get
        Set(value As String)
            Me._HelpFile = value
        End Set
    End Property

    Private _InitialRead As Boolean
    Private Property InitialRead As Boolean
        Get
            Return Me._InitialRead
        End Get
        Set(value As Boolean)
            Me._InitialRead = value
        End Set
    End Property

    Private _Buffers As List(Of String)
    Private Property Buffers As List(Of String)
        Get
            If Me._Buffers Is Nothing Then
                Me._Buffers = New List(Of String)
            End If
            Return Me._Buffers
        End Get
        Set(value As List(Of String))
            Me._Buffers = value
        End Set
    End Property

    Private _BufferIndex As Integer
    Private Property BufferIndex As Integer
        Get
            Return Me._BufferIndex
        End Get
        Set(value As Integer)
            Me._BufferIndex = value
        End Set
    End Property

    Private _CurrentLine As String
    Private Property CurrentLine As String
        Get
            Return Me._CurrentLine
        End Get
        Set(value As String)
            Me._CurrentLine = value
        End Set
    End Property

    Private _CurrentLineCount As Integer
    Private Property CurrentLineCount As Integer
        Get
            Return Me._CurrentLineCount
        End Get
        Set(value As Integer)
            Me._CurrentLineCount = value
        End Set
    End Property

    Private ReadOnly Property Page As Integer
        Get
            Return (Me.BufferIndex + 1)
        End Get
    End Property

    Private _LastPage As Integer
    Private Property LastPage As Integer
        Get
            Return Me._LastPage
        End Get
        Set(value As Integer)
            Me._LastPage = value
        End Set
    End Property

    Private ReadOnly Property BeginPage As Integer
        Get
            Return 1
        End Get
    End Property

    Private _Helps As List(Of String)
    Private Property Helps As List(Of String)
        Get
            If Me._Helps Is Nothing Then
                Me._Helps = New List(Of String)
            End If
            Return Me._Helps
        End Get
        Set(value As List(Of String))
            Me._Helps = value
        End Set
    End Property

    Private _ExitFlag As Boolean
    Private Property ExitFlag As Boolean
        Get
            Return Me._ExitFlag
        End Get
        Set(value As Boolean)
            Me._ExitFlag = value
        End Set
    End Property

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim hs As List(Of String) = Directory.GetFiles(RpaCui.SystemDllDirectory).ToList
        Me.Helps = hs.FindAll(
                        Function(h)
                            If Path.GetExtension(h) = ".txt" Then
                                Return True
                            Else
                                Return False
                            End If
                        End Function
                   )
        Me.Helps.Add("終了する")

        Do
            Console.WriteLine($"id  ファイル名")
            Console.WriteLine($"______________________________________________________________________")
            For Each help In Me.Helps
                Dim idxno As String = String.Format("{0, 3}", Me.Helps.IndexOf(help))
                Dim helpname As String = Path.GetFileName(help)
                Console.WriteLine($"{idxno} {helpname}")
            Next
            Console.WriteLine()

            Dim idx As Integer = -1
            Do
                idx = -1
                Console.WriteLine($"確認したいヘルプの 'id' を指定してください")
                Dim idxstr As String = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()

                If Not IsNumeric(idxstr) Then
                    Continue Do
                End If

                idx = Integer.Parse(idxstr)
                If idx < Me.Helps.IndexOf(Me.Helps.First) Then
                    Continue Do
                End If
                If idx > Me.Helps.IndexOf(Me.Helps.Last) Then
                    Continue Do
                End If
                Exit Do
            Loop Until False

            If idx = Me.Helps.IndexOf(Me.Helps.Last) Then
                Exit Do
            End If

            Me.HelpFile = Me.Helps(idx)
            Dim i As Integer = Read(dat)
        Loop Until False

        Return 0
    End Function

    Private Function Read(ByRef dat As RpaDataWrapper) As Integer
        Me.Buffers = Nothing
        Me.BufferIndex = 0
        Me.Reader = New StreamReader(Me.HelpFile, Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING))

        Dim i As Integer = ReadBuffers()

        Call Console.Clear()
        Dim j As Integer = GoForward()
        Do
            Dim k As ConsoleKey = GetPageNavigationKey()

            Call Console.Clear()

            If Me.ForwardChars.Contains(k) Then
                Dim h As Integer = GoForward()
            End If
            If Me.BackChars.Contains(k) Then
                Dim h As Integer = GoBack()
            End If

            If Me.ExitFlag Then
                Exit Do
            End If
        Loop Until False

        Me.Reader.Close()
        Me.Reader.Dispose()

        Return 0
    End Function

    Private Function GetPageNavigationKey() As ConsoleKey
        Dim k As ConsoleKeyInfo
        Do
            k = Nothing
            Console.WriteLine(Me.Navigater)
            k = Console.ReadKey(True)
            If Me.ForwardChars.Contains(k.Key) Or Me.BackChars.Contains(k.Key) Then
                Exit Do
            End If
        Loop Until False

        Return k.Key
    End Function

    ' 現状、１行がウィンドウ行数より多くなる場合は想定していない。無限ループするかも・・・
    Private Function ReadBuffers() As Integer
        Dim pagebuf As String = vbNullString
        Dim linesum As Integer = 0
        Do
            Dim [line] As String = Me.Reader.ReadLine()
            Dim linelen As Integer = Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING).GetByteCount([line])
            Dim linecnt As Integer = IIf(linelen < Console.WindowWidth, 1, ((linelen / Console.WindowWidth) + 1))

            If Me.Reader.EndOfStream Or (linesum + linecnt) >= (Console.WindowHeight - 3) Then
                Dim lineres As Integer = (Console.WindowHeight - 2) - linesum
                pagebuf &= Strings.StrDup(lineres, vbCrLf)

                Me.Buffers.Add(pagebuf)
                pagebuf = vbNullString
                linesum = 0
            End If

            If Me.Reader.EndOfStream Then
                Exit Do
            End If

            pagebuf &= $"{[line]}{vbCrLf}"
            linesum += linecnt

        Loop Until False

        Me.LastPage = Me.Buffers.Count + 1

        Return 0
    End Function

    Private Function GoForward() As Integer
        Me.BufferIndex += 1
        Console.WriteLine(Me.Buffers(Me.BufferIndex))
        Return 0
    End Function

    Private Function GoBack() As Integer
        Me.BufferIndex -= 1
        Console.WriteLine(Me.Buffers(Me.BufferIndex))
        Return 0
    End Function

    Sub New()
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
