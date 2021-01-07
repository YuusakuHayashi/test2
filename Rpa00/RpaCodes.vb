Imports System.Linq
Imports System.Reflection

Public Module RpaCodes
    Public ReadOnly Property RpaObject(rpa As Object) As Object
        Get
            Dim asm As Assembly
            Dim [mod] As [Module]
            Dim [type] As Object
            Dim dpath As String = RpaModule.GetFileText($"{RpaCui.SystemDirectory}\debugpath")
            Select Case rpa.RobotName
                Case "rpa01"
                    If String.IsNullOrEmpty(dpath) Then
                        asm = Assembly.LoadFrom($"{RpaCui.SystemDllDirectory}\Rpa01.dll")
                    Else
                        asm = Assembly.LoadFrom($"{dpath}\Rpa01\obj\Debug\Rpa01.dll")
                    End If
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
