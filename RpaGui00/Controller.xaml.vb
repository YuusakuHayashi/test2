Imports System.Reflection
Imports System.IO

Public Class Controller
    ' !!! RpaCuiと同じプロパティ。同期して変更する必要あり
    Public Shared ReadOnly Property SystemDirectory As String
        Get
            Dim [dir] As String = System.Environment.GetEnvironmentVariable("USERPROFILE") & "\rpa_project"
            If Not Directory.Exists([dir]) Then
                Directory.CreateDirectory([dir])
            End If
            Return [dir]
        End Get
    End Property

    ' !!! RpaCuiと同じプロパティ。同期して変更する必要あり
    Private Shared _SystemDllDirectory As String
    Public Shared Property SystemDllDirectory As String
        Get
            If String.IsNullOrEmpty(Controller._SystemDllDirectory) Then
                Controller._SystemDllDirectory = $"{Controller.SystemDirectory}\dll"
                If Not Directory.Exists(Controller._SystemDllDirectory) Then
                    Directory.CreateDirectory(Controller._SystemDllDirectory)
                End If
            End If
            Return Controller._SystemDllDirectory
        End Get
        Set(value As String)
            Controller._SystemDllDirectory = value
        End Set
    End Property

    ' !!! RpaCuiと同じプロパティ。同期して変更する必要あり
    Public Shared ReadOnly Property SystemIniFileName As String
        Get
            Return $"{Controller.SystemDirectory}\rpa.ini"
        End Get
    End Property

    Sub New()
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。        
        Dim vm As ControllerViewModel
        ' デバッグする場合
        'Controller.SystemDllDirectory = $"\\Coral\個人情報-林祐\project\wpf\test2\RpaCui\debugrobot"
        'Controller.SystemDllDirectory = $"C:\Users\yuusa\project\test2\RpaCui\debugrobot"

        '-----------------------------------------------------------------------------------------'
        ' !!! RpaCuiと同じロジック。同期して変更する必要あり？
        ' DLLロード
        ' アップデート・デバッグパスには現状対応してない
        Dim rpa00dll As String = $"{Controller.SystemDllDirectory}\Rpa00.dll"
        Dim asm As Assembly = Assembly.LoadFrom(rpa00dll)
        Dim [mod] As [Module] = asm.GetModule("Rpa00.dll")
        Dim dat_type = [mod].GetType("Rpa00.RpaDataWrapper")
        Dim dat = Activator.CreateInstance(dat_type)
        ' セーブ情報読み込み
        If Not File.Exists(Controller.SystemIniFileName) Then
            dat.Initializer.Save(Controller.SystemIniFileName, dat.Initializer)
        End If
        dat.Initializer = dat.Initializer.Load(Controller.SystemIniFileName)
        dat.Project = dat.System.LoadCurrentRpa(dat)
        '-----------------------------------------------------------------------------------------'        
        vm = New ControllerViewModel
        vm.Data = dat
        vm.Initialize()

        Me.DataContext = vm
    End Sub
End Class
