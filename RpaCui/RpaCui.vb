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

    Public ReadOnly Property SystemUpdateDllDirectory As String
        Get
            Dim [dir] As String = $"{RpaCui.SystemDirectory}\updatedll"
            If Not Directory.Exists([dir]) Then
                Directory.CreateDirectory([dir])
            End If
            Return [dir]
        End Get
    End Property

    Public ReadOnly Property SystemPointUpdateDllDirectory As String
        Get
            Dim [dir] As String = $"{RpaCui.SystemDirectory}\pointupdatedll"
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

            If Not Directory.GetFiles(RpaCui.SystemUpdateDllDirectory).Count = 0 Then
                Dim srcs1 As List(Of String) = Directory.GetFiles(RpaCui.SystemUpdateDllDirectory).ToList
                For Each src In srcs1
                    Dim dst As String = $"{RpaCui.SystemDllDirectory}\{Path.GetFileName(src)}"
                    Console.WriteLine($"アップグレードを適用しています... Update --> {dst}")
                    File.Copy(src, dst, True)
                    File.Delete(src)
                Next
            End If

            If Not Directory.GetFiles(RpaCui.SystemPointUpdateDllDirectory).Count = 0 Then
                Dim srcs2 As List(Of String) = Directory.GetFiles(RpaCui.SystemPointUpdateDllDirectory).ToList

                ' 削除のため差集合的なものを求める
                Dim dsts As List(Of String) = Directory.GetFiles(RpaCui.SystemDllDirectory).ToList
                Dim srcs3 As New List(Of String)
                For Each src In srcs2
                    srcs3.Add(Path.GetFileName(src))
                Next
                Dim subs As List(Of String) = dsts.FindAll(Function(d)
                                                               Dim b As Boolean = False
                                                               If Path.GetExtension(d) = ".dll" Then
                                                                   If Not srcs3.Contains(Path.GetFileName(d)) Then
                                                                       b = True
                                                                   End If
                                                               End If
                                                               Return b
                                                           End Function)
                For Each [sub] In subs
                    Console.WriteLine($"特定時点へのアップグレード／ダウングレードを適用しています... Delete --> {[sub]}")
                    File.Delete([sub])
                Next

                For Each src In srcs2
                    Dim dst As String = $"{RpaCui.SystemDllDirectory}\{Path.GetFileName(src)}"
                    Console.WriteLine($"特定時点へのアップグレード／ダウングレードを適用しています... Update --> {dst}")
                    File.Copy(src, dst, True)
                    File.Delete(src)
                Next
            End If

            ' DLLロード
            Dim asm As Assembly = Assembly.LoadFrom(rpa00dll)
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
