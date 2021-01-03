Imports System.Runtime.InteropServices
Imports System.IO
Imports Rpa00

Public Class RpaMacroUtility : Inherits RpaUtilityBase
    Public Overrides ReadOnly Property ExecuteHandler(dat As Object) As Object
        Get
            Dim cmd As Object
            Select Case dat.Transaction.MainCommand
                Case "UpdateMacro" : cmd = New UpdateMacroCommand(Me)
                Case Else : cmd = Nothing
            End Select
            Return cmd
        End Get
    End Property

    Private ReadOnly Property _MacroUtilityDirectory As String
        Get
            Dim d = $"{CommonProject.SystemDirectory}\MacroUtility"
            If Not Directory.Exists(d) Then
                Directory.CreateDirectory(d)
            End If
            Return d
        End Get
    End Property

    Private ReadOnly Property _MacroFileName As String
        Get
            Dim f = $"{Me._MacroUtilityDirectory}\macro.xlsm"
            Return f
        End Get
    End Property

    Public Function InvokeMacro(ByVal method As String, args() As Object)
        Dim exapp, exbooks, exbook
        Dim macrofile As String

        Try
            exapp = CreateObject("Excel.Application")
            Try
                exbooks = exapp.Workbooks
                Try
                    exbook = exbooks.Open(Me._MacroFileName, 0)
                    macrofile = Path.GetFileName(Me._MacroFileName)
                    exapp.Run($"{macrofile}!{method}", args)
                Finally
                    If exbook IsNot Nothing Then
                        exbook.Close()
                    End If
                    Marshal.ReleaseComObject(exbook)
                End Try
            Finally
                Marshal.ReleaseComObject(exbooks)
            End Try
            Return 0
        Finally
            If exapp IsNot Nothing Then
                exapp.Quit()
            End If
            Marshal.ReleaseComObject(exapp)
        End Try
    End Function

    Private Class UpdateMacroCommand : Inherits RpaUtilityCommandBase
        Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
            Get
                Return {1, 999}
            End Get
        End Property

        Public Overrides Function CanExecute(ByRef dat As Object) As Boolean
            If Not File.Exists(Me.Parent._MacroFileName) Then
                Console.WriteLine($"マクロファイル '{Me.Parent._MacroFileName}' が存在しません")
                Console.WriteLine($"開発者から入手してください")
                Console.WriteLine()
                Return False
            End If
            Return True
        End Function

        Private Parent As RpaMacroUtility
        Public Overrides Function Execute(ByRef dat As Object) As Integer
            For Each p In dat.Transaction.Parameters
                Dim bas As String = $"{Parent._MacroUtilityDirectory}\{p}"
                If File.Exists(bas) Then
                    Console.WriteLine($"指定マクロ '{p}' をインストールします")
                    Call Parent.InvokeMacro("MacroImporter.Main", {bas})
                    Console.WriteLine()
                Else
                    Console.WriteLine($"指定マクロ '{p}' は存在しません")
                    Console.WriteLine()
                End If
            Next
            Return 0
        End Function

        Sub New(p As RpaMacroUtility)
            Me.Parent = p
        End Sub
    End Class
End Class
