Imports System.Linq
Imports System.Reflection

Module RpaCodes
    Public ReadOnly Property RpaObject(ByVal code As String) As Object
        Get
            Dim asm As Assembly
            Dim [mod] As [Module]
            Dim [type] As Object
            Select Case code
                Case "rpa01"
                    asm = Assembly.LoadFrom("C:\Users\yuusa\project\test2\Rpa01\obj\Debug\Rpa01.dll")
                    [mod] = asm.GetModule("Rpa01.dll")
                    [type] = [mod].GetType("Rpa01.Rpa01")
                    If [type] IsNot Nothing Then
                        Return Activator.CreateInstance([type])
                    End If
                Case Else
                    Return Nothing
            End Select
        End Get
    End Property
End Module
