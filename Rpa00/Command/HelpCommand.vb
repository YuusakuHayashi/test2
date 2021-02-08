Imports System.IO

Public Class HelpCommand : Inherits RpaCommandBase

    Private Function Check(ByRef dat As RpaDataWrapper) As Boolean
        Return True
    End Function

    Private ReadOnly Property Navigater As String
        Get
            Dim pstr As String = $"p. {Me.Page.ToString}/{Me.LastPage}"
            Dim pstrlen As Integer = pstr.Length
            Dim bstr As String = IIf(Me.Page = Me.BeginPage, "    ", "  <<")
            Dim bstrlen As Integer = bstr.Length
            Dim fstr As String = IIf(Me.Page = Me.LastPage, "END ", ">> ")
            Dim fstrlen As Integer = fstr.Length

            Dim spclen As Integer = (Console.WindowWidth - 1) - (pstrlen + bstrlen + fstrlen)
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
                Me._ForwardChars.Add(ConsoleKey.DownArrow)
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
                Me._BackChars.Add(ConsoleKey.Subtract)
                Me._BackChars.Add(ConsoleKey.UpArrow)
            End If
            Return Me._BackChars
        End Get
    End Property

    Private _EndChars As List(Of ConsoleKey)
    Private ReadOnly Property EndChars As List(Of ConsoleKey)
        Get
            If Me._EndChars Is Nothing Then
                Me._EndChars = New List(Of ConsoleKey)
                Me._EndChars.Add(ConsoleKey.End)
            End If
            Return Me._EndChars
        End Get
    End Property

    Private _TopChars As List(Of ConsoleKey)
    Private ReadOnly Property TopChars As List(Of ConsoleKey)
        Get
            If Me._TopChars Is Nothing Then
                Me._TopChars = New List(Of ConsoleKey)
                Me._TopChars.Add(ConsoleKey.Home)
            End If
            Return Me._TopChars
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

    Private _OldHeight As Integer
    Private Property OldHeight As Integer
        Get
            Return Me._OldHeight
        End Get
        Set(value As Integer)
            Me._OldHeight = value
        End Set
    End Property

    Private _OldWidth As Integer
    Private Property OldWidth As Integer
        Get
            Return Me._OldWidth
        End Get
        Set(value As Integer)
            Me._OldWidth = value
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
            Console.WriteLine($" ID  | ファイル名")
            Console.WriteLine($"-----+----------------------------------------------------------------")
            For Each help In Me.Helps
                Dim idxno As String = String.Format("{0, 4}", Me.Helps.IndexOf(help))
                Dim helpname As String = Path.GetFileName(help)
                Console.WriteLine($"{idxno} | {helpname}")
            Next
            Console.WriteLine()

            Dim idx As Integer = -1
            Do
                idx = -1
                Console.WriteLine($"確認したいヘルプの 'ID' を指定してください")
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
        Me.BufferIndex = 0

        Do
            Me.OldHeight = Console.WindowHeight
            Me.OldWidth = Console.WindowWidth

            Me.Buffers = Nothing
            Me.Reader = New StreamReader(Me.HelpFile, Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING))

            Dim i As Integer = ReadBuffers()
            Do
                If (Me.Page < Me.BeginPage) Or (Me.Page > Me.LastPage) Or (Console.WindowHeight <> Me.OldHeight) Or (Console.WindowWidth <> Me.OldWidth) Then
                    Exit Do
                End If

                Call Console.Clear()
                Console.WriteLine(Me.Buffers(Me.BufferIndex))

                Dim k As ConsoleKey = GetPageNavigationKey()

                If Me.ForwardChars.Contains(k) Then
                    Dim h As Integer = GoForward()
                End If
                If Me.BackChars.Contains(k) Then
                    Dim h As Integer = GoBack()
                End If
                If Me.TopChars.Contains(k) Then
                    Dim h As Integer = GoTop()
                End If
                If Me.EndChars.Contains(k) Then
                    Dim h As Integer = GoEnd()
                End If
            Loop Until False

            Call Console.Clear()
            Me.Reader.Close()
            Me.Reader.Dispose()

            If (Me.Page < Me.BeginPage) Or (Me.Page > Me.LastPage) Then
                Exit Do
            End If
        Loop Until False

        Return 0
    End Function

    Private Function GetPageNavigationKey() As ConsoleKey
        Dim k As ConsoleKeyInfo
        Do
            k = Nothing

            Console.Write(Me.Navigater)
            k = Console.ReadKey(True)

            If Me.ForwardChars.Contains(k.Key) Or Me.BackChars.Contains(k.Key) Or Me.TopChars.Contains(k.Key) Or Me.EndChars.Contains(k.Key) Then
                Exit Do
            End If

            Call Console.Clear()
            Console.WriteLine(Me.Buffers(Me.BufferIndex))
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

        Me.LastPage = Me.Buffers.Count

        Return 0
    End Function

    Private Function GoForward() As Integer
        Me.BufferIndex += 1
        Return 0
    End Function

    Private Function GoBack() As Integer
        Me.BufferIndex -= 1
        Return 0
    End Function

    Private Function GoTop() As Integer
        Me.BufferIndex = 0
        Return 0
    End Function

    Private Function GoEnd() As Integer
        Me.BufferIndex = Me.Buffers.IndexOf(Me.Buffers.Last)
        Return 0
    End Function

    Sub New()
        Me.CanExecuteHandler = AddressOf Check
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
