Imports System.Linq
Imports System.Reflection
Imports System.IO

Public Module RpaCodes
    ' プラグインを追加する場合には、ここに記述
    Public Function CreateRobotObject(ByRef rpa As Object) As Object
        ' Test
        Dim dllpath As String = $"\\Coral\個人情報-林祐\project\wpf\test2\RpaCui\debugrobot\{rpa.RobotName}.dll"
        'Dim dllpath As String = $"{RpaCui.SystemDllDirectory}\{rpa.RobotName}.dll"

        If Not File.Exists(dllpath) Then
            Return Nothing
        End If

        Try
            Dim modname As String = Path.GetFileName(dllpath)
            Dim prjname As String = Path.GetFileNameWithoutExtension(dllpath)
            Dim clsname As String = Path.GetFileNameWithoutExtension(dllpath)

            Dim asm As Assembly = Assembly.LoadFrom(dllpath)
            Dim [mod] As [Module] = asm.GetModule(modname)
            Dim cls As Object = [mod].GetType($"{prjname}.{clsname}")
            Dim obj As Object = Activator.CreateInstance(cls)

            Return obj
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    'Public ReadOnly Property RobotObject(rpa As Object) As Object
    '    Get
    '        Dim asm As Assembly
    '        Dim [mod] As [Module]
    '        Dim [type] As Object
    '        Dim dllpath As String = vbNullString
    '        Select Case rpa.RobotName
    '            Case "Rpa01"
    '                dllpath = $"{RpaCui.SystemDllDirectory}\Rpa01.dll"
    '                If File.Exists(dllpath) Then
    '                    asm = Assembly.LoadFrom(dllpath)
    '                    [mod] = asm.GetModule("Rpa01.dll")
    '                    [type] = [mod].GetType("Rpa01.Rpa01")
    '                    Return Activator.CreateInstance([type])
    '                Else
    '                    Return Nothing
    '                End If
    '            Case "Rpa02"
    '                dllpath = $"{RpaCui.SystemDllDirectory}\Rpa02.dll"
    '                If File.Exists(dllpath) Then
    '                    asm = Assembly.LoadFrom(dllpath)
    '                    [mod] = asm.GetModule("Rpa02.dll")
    '                    [type] = [mod].GetType("Rpa02.Rpa02")
    '                    Return Activator.CreateInstance([type])
    '                Else
    '                    Return Nothing
    '                End If
    '            Case "Rpa07"
    '                dllpath = $"{RpaCui.SystemDllDirectory}\Rpa07.dll"
    '                If File.Exists(dllpath) Then
    '                    asm = Assembly.LoadFrom(dllpath)
    '                    [mod] = asm.GetModule("Rpa07.dll")
    '                    [type] = [mod].GetType("Rpa07.Rpa07")
    '                    Return Activator.CreateInstance([type])
    '                Else
    '                    Return Nothing
    '                End If
    '            Case Else
    '                Return Nothing
    '        End Select
    '    End Get
    'End Property


    Public ReadOnly Property RpaUtilityObject(util As String) As Object
        Get
            Select Case util
                Case "macroutility"
                    Return (New RpaMacroUtility)
                Case "printerutility"
                    Return (New RpaPrinterUtility)
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property
End Module
