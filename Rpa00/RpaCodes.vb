Imports System.Linq
Imports System.Reflection

Public Module RpaCodes
    Public ReadOnly Property RpaObject(rpa As Object) As Object
        Get
            Dim asm As Assembly
            Dim [mod] As [Module]
            Dim [type] As Object
            Select Case rpa.RobotName
                Case "rpa01"
                    'asm = Assembly.LoadFrom("C:\Users\yuusa\project\test2\Rpa01\obj\Debug\Rpa01.dll")
                    asm = Assembly.LoadFrom("\\Coral\個人情報-林祐\project\wpf\test2\Rpa01\obj\Debug\Rpa01.dll")
                    'asm = Assembly.LoadFrom($"{CommonProject.SystemDllDirectory}\Rpa01.dll")
                    [mod] = asm.GetModule("Rpa01.dll")
                    [type] = [mod].GetType("Rpa01.Rpa01")
                    If [type] IsNot Nothing Then
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
