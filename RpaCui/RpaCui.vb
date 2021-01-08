Imports System.Reflection
Imports System.IO

Public Module RpaCui
    Public ReadOnly Property SystemDirectory As String
        Get
            Dim [dir] As String = System.Environment.GetEnvironmentVariable("USERPROFILE") & "\rpa_project"
            If Not Directory.Exists([dir]) Then
                Directory.CreateDirectory([dir])
            End If
            Return [dir]
        End Get
    End Property

    Public ReadOnly Property SystemDllDirectory As String
        Get
            Dim [dir] As String = $"{RpaCui.SystemDirectory}\dll"
            If Not Directory.Exists([dir]) Then
                Directory.CreateDirectory([dir])
            End If
            Return [dir]
        End Get
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

    Public ReadOnly Property SystemIniFileName As String
        Get
            Return $"{RpaCui.SystemDirectory}\rpa.ini"
        End Get
    End Property

    Public Sub Main(ByVal args As String())
        Dim ptype As String = vbNullString
        Dim txt As String = vbNullString
        Dim inv As Boolean = False
        Dim rpa00dll As String = $"{RpaCui.SystemDllDirectory}\Rpa00.dll"

        Try
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
            'Dim asm As Assembly = Assembly.LoadFrom(rpa00dll)
            Dim asm As Assembly = Assembly.LoadFrom("\\Coral\個人情報-林祐\project\wpf\test2\Rpa00\bin\Debug\Rpa00.dll")
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
        Finally
        End Try
    End Sub
End Module
