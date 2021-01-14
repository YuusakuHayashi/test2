Imports System.Reflection
Imports System.IO

Public Module RpaCui
    Private _DebugMode As Boolean
    Private Property DebugMode As Boolean
        Get
            Return RpaCui._DebugMode
        End Get
        Set(value As Boolean)
            RpaCui._DebugMode = value
        End Set
    End Property

    Public ReadOnly Property SystemDirectory As String
        Get
            Dim [dir] As String = System.Environment.GetEnvironmentVariable("USERPROFILE") & "\rpa_project"
            If Not Directory.Exists([dir]) Then
                Directory.CreateDirectory([dir])
            End If
            Return [dir]
        End Get
    End Property

    'Public ReadOnly Property SystemDllDirectory As String
    '    Get
    '        Dim [dir] As String = $"{RpaCui.SystemDirectory}\dll"
    '        If Not Directory.Exists([dir]) Then
    '            Directory.CreateDirectory([dir])
    '        End If
    '        Return [dir]
    '    End Get
    'End Property

    Private _SystemDllDirectory As String
    Public Property SystemDllDirectory As String
        Get
            If String.IsNullOrEmpty(RpaCui._SystemDllDirectory) Then
                RpaCui._SystemDllDirectory = $"{RpaCui.SystemDirectory}\dll"
                If Not Directory.Exists(RpaCui._SystemDllDirectory) Then
                    Directory.CreateDirectory(RpaCui._SystemDllDirectory)
                End If
            End If
            Return RpaCui._SystemDllDirectory
        End Get
        Set(value As String)
            RpaCui._SystemDllDirectory = value
        End Set
    End Property

    'Private _SystemDllDirectory As String
    'Public Property SystemDllDirectory As String
    '    Get
    '        If String.IsNullOrEmpty(RpaCui._SystemDllDirectory) Then
    '            RpaCui._SystemDllDirectory = $"{RpaCui.SystemDirectory}\dll"
    '        End If
    '        If Not Directory.Exists(RpaCui._SystemDllDirectory) Then
    '            Directory.CreateDirectory(RpaCui._SystemDllDirectory)
    '        End If
    '        Return RpaCui._SystemDllDirectory
    '    End Get
    '    Set(value As String)
    '        RpaCui._SystemDllDirectory = value
    '    End Set
    'End Property

    Public ReadOnly Property SystemUpdateDllDirectory As String
        Get
            Dim [dir] As String = $"{RpaCui.SystemDirectory}\updatedll"
            If Not Directory.Exists([dir]) Then
                Directory.CreateDirectory([dir])
            End If
            Return [dir]
        End Get
    End Property

    Public ReadOnly Property SystemIniFileName As String
        Get
            Return $"{RpaCui.SystemDirectory}\rpa.ini"
        End Get
    End Property

    Public Sub Main(ByVal args As String())
        Try
            If args.Count > 0 Then
                RpaCui.SystemDllDirectory = args(0)
            End If

            Dim rpa00dll As String = $"{RpaCui.SystemDllDirectory}\Rpa00.dll"

            ' アップデート適用
            Do
                If Not Directory.GetFiles(RpaCui.SystemUpdateDllDirectory).Count = 0 Then
                    Dim updatedll As String = Directory.GetFiles(RpaCui.SystemUpdateDllDirectory)(0)
                    Dim dst As String = $"{RpaCui.SystemDllDirectory}\{Path.GetFileName(updatedll)}"
                    Console.WriteLine($"アップデートを適用しています...{dst}")
                    File.Copy(updatedll, dst, True)
                    File.Delete(updatedll)
                Else
                    Exit Do
                End If
            Loop Until False

            ' DLLロード
            Dim asm As Assembly = Assembly.LoadFrom(rpa00dll)
            'Dim asm As Assembly = Assembly.LoadFrom("C:\Users\yuusa\project\test2\Rpa00\bin\Debug\Rpa00.dll")
            Dim [mod] As [Module] = asm.GetModule("Rpa00.dll")
            Dim dat_type = [mod].GetType("Rpa00.RpaDataWrapper")
            Dim dat = Activator.CreateInstance(dat_type)

            ' セーブ情報読み込み
            If Not File.Exists(RpaCui.SystemIniFileName) Then
                dat.Initializer.Save(RpaCui.SystemIniFileName, dat.Initializer)
            End If
            dat.Initializer = dat.Initializer.Load(RpaCui.SystemIniFileName)
            dat.Project = dat.System.LoadCurrentRpa(dat)

            ' 実行
            Do Until dat.Transaction.ExitFlag
                Call dat.System.Main(dat)
            Loop
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Console.WriteLine("RpaCui.exe は異常終了しました")
            Console.ReadLine()
        Finally
        End Try
    End Sub
End Module
