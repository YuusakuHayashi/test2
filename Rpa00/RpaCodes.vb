Imports System.Linq
Imports System.Reflection
Imports System.IO

Public Module RpaCodes

    ' プラグインを追加する場合には、ここに記述
    Public ReadOnly Property RpaObject(rpa As Object) As Object
        Get
            Dim asm As Assembly
            Dim [mod] As [Module]
            Dim [type] As Object
            Dim dllpath As String = vbNullString
            Select Case rpa.RobotName
                Case "rpa01"
                    'dllpath = $"{RpaCui.SystemDllDirectory}\Rpa01.dll"
                    dllpath = "\\Coral\個人情報-林祐\project\wpf\test2\Rpa01\bin\Debug\Rpa01.dll"
                    If File.Exists(dllpath) Then
                        asm = Assembly.LoadFrom(dllpath)
                        [mod] = asm.GetModule("Rpa01.dll")
                        [type] = [mod].GetType("Rpa01.Rpa01")
                        Return Activator.CreateInstance([type])
                    Else
                        Return Nothing
                    End If
                Case "rpa07"
                    'dllpath = $"{RpaCui.SystemDllDirectory}\Rpa01.dll"
                    dllpath = "\\Coral\個人情報-林祐\project\wpf\test2\Rpa07\bin\Debug\Rpa07.dll"
                    If File.Exists(dllpath) Then
                        asm = Assembly.LoadFrom(dllpath)
                        [mod] = asm.GetModule("Rpa07.dll")
                        [type] = [mod].GetType("Rpa07.Rpa07")
                        Return Activator.CreateInstance([type])
                    Else
                        Return Nothing
                    End If
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property


    Public ReadOnly Property RpaUtilityObject(util As String) As Object
        Get
            Select Case util
                Case "MacroUtility"
                    Return (New RpaMacroUtility)
                Case "PrinterUtility"
                    Return (New RpaPrinterUtility)
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property
End Module
