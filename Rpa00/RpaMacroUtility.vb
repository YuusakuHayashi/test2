Imports System.Runtime.InteropServices
Imports System.IO
Imports Rpa00

Public Class RpaMacroUtility : Inherits RpaUtilityBase
    Public Overrides ReadOnly Property CommandHandler(dat As Object) As Object
        Get
            Dim cmd As Object
            Select Case dat.Transaction.MainCommand
                Case "update" : cmd = New UpdateMacroCommand(Me)
                Case "check" : cmd = New CheckMacroCommand(Me)
                Case Else : cmd = Nothing
            End Select
            Return cmd
        End Get
    End Property

    Public ReadOnly Property MacroUtilityDirectory As String
        Get
            Dim d = $"{CommonProject.SystemDirectory}\MacroUtility"
            If Not Directory.Exists(d) Then
                Directory.CreateDirectory(d)
            End If
            Return d
        End Get
    End Property

    Public ReadOnly Property MacroFileName As String
        Get
            Dim f = $"{Me.MacroUtilityDirectory}\macro.xlsm"
            Return f
        End Get
    End Property

    Public Function InvokeMacro(ByVal method As String, args() As Object) As Integer
        Dim rtn As Integer
        Dim exapp, exbooks, exbook
        Dim macrofile As String

        rtn = -1
        Try
            exapp = CreateObject("Excel.Application")
            Try
                exbooks = exapp.Workbooks
                Try
                    exbook = exbooks.Open(Me.MacroFileName, 0)
                    macrofile = Path.GetFileName(Me.MacroFileName)
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
            rtn = 0
        Finally
            If exapp IsNot Nothing Then
                exapp.Quit()
            End If
            Marshal.ReleaseComObject(exapp)
        End Try
        Return rtn
    End Function

    Public Function CallMacro(ByVal method As String, args() As Object) As Object
        Dim exapp, exbooks, exbook
        Dim macrofile As String
        Dim obj As Object = Nothing

        Try
            exapp = CreateObject("Excel.Application")
            Try
                exbooks = exapp.Workbooks
                Try
                    exbook = exbooks.Open(Me.MacroFileName, 0)
                    macrofile = Path.GetFileName(Me.MacroFileName)
                    obj = exapp.Run($"{macrofile}!{method}", args)
                Finally
                    If exbook IsNot Nothing Then
                        exbook.Close()
                    End If
                    Marshal.ReleaseComObject(exbook)
                End Try
            Finally
                Marshal.ReleaseComObject(exbooks)
            End Try
        Finally
            If exapp IsNot Nothing Then
                exapp.Quit()
            End If
            Marshal.ReleaseComObject(exapp)
        End Try
        Return obj
    End Function

    Private Class UpdateMacroCommand : Inherits RpaUtilityCommandBase
        Public Overrides ReadOnly Property ExecutableParameterCount As Integer()
            Get
                Return {1, 999}
            End Get
        End Property

        Public Overrides Function CanExecute(ByRef dat As Object) As Boolean
            If Not File.Exists(Me.Parent.MacroFileName) Then
                Console.WriteLine($"マクロファイル '{Me.Parent.MacroFileName}' が存在しません")
                Console.WriteLine($"開発者から入手してください")
                Console.WriteLine()
                Return False
            End If
            Return True
        End Function

        Private Parent As RpaMacroUtility
        Public Overrides Function Execute(ByRef dat As Object) As Integer
            For Each p In dat.Transaction.Parameters
                Dim bas As String = $"{Parent.MacroUtilityDirectory}\{p}"
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
