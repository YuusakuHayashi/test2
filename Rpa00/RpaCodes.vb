Imports System.Linq
Imports System.Reflection

Public Module RpaCodes
    'Public Enum ProjectArchitecture
    '    IntranetClientServer = 1
    '    StandAlone = 2
    '    ClientServer = 3
    'End Enum

    'Public Enum ProjectUserLevel
    '    User = 1
    '    HiLevelUser = 2
    'End Enum

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


    ' ユーティリティの変換登録
    'Public ReadOnly Property RpaUtility(util As String) As RpaUtility
    '    Get
    '        Const MACROUTILITY As String = "MacroUtility"
    '        Const PRINTERUTILITY As String = "PrinterUtility"
    '        Select Case util
    '            Case MACROUTILITY
    '                Return New RpaUtility With {
    '                    .UtilityName = MACROUTILITY,
    '                    .UtilityObject = (New RpaMacroUtility)
    '                }
    '            Case PRINTERUTILITY
    '                Return New RpaUtility With {
    '                    .UtilityName = PRINTERUTILITY,
    '                    .UtilityObject = (New RpaPrinterUtility)
    '                }
    '            Case Else
    '                Return Nothing
    '        End Select
    '    End Get
    'End Property

    'Public ReadOnly Property RpaUtility(util As String) As Object
    '    Get
    '        Dim uobj = RpaUtilityObject(util)
    '        If uobj Is Nothing Then
    '            Return Nothing
    '        Else
    '            Return (New RpaUtility With {
    '                .UtilityName = util,
    '                .UtilityObject = uobj
    '            })
    '        End If
    '    End Get
    'End Property

    Public ReadOnly Property RpaUtilityObject(util As String) As Object
        Get
            Const MACROUTILITY As String = "MacroUtility"
            Const PRINTERUTILITY As String = "PrinterUtility"
            Select Case util
                Case MACROUTILITY
                    Return (New RpaMacroUtility)
                Case PRINTERUTILITY
                    Return (New RpaPrinterUtility)
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property

End Module
