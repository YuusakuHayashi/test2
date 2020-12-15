Imports System.Runtime.InteropServices
Imports System.IO

Public Class RpaMacroUtility
    Private _MacroFileName As String
    Public Property MacroFileName As String
        Get
            Return Me._MacroFileName
        End Get
        Set(value As String)
            Me._MacroFileName = value
        End Set
    End Property

    Public Sub InvokeMacro(ByVal method As String, args() As Object)
        Dim exapp, exbooks, exbook
        Dim macrofile As String
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
        Finally
            If exapp IsNot Nothing Then
                exapp.Quit()
            End If
            Marshal.ReleaseComObject(exapp)
        End Try
    End Sub

    ' マクロの更新
    '---------------------------------------------------------------------------------------------'
    Private Function UpdateMacro(ByRef trn As RpaTransaction, ByRef rpa As RpaProject) As Integer
        Dim bas As String
        If trn.Parameters.Count = 0 Then
            Console.WriteLine("パラメータが指定されていません: " & trn.CommandText)
            Return 1000
        End If

        For Each p In trn.Parameters
            bas = RpaProject.SYSTEM_SCRIPT_DIRECTORY & "\" & p
            If File.Exists(bas) Then
                Console.WriteLine($"指定マクロ '{p}' をインストールします")
                Call InvokeMacro("MacroImporter.Main", {bas})
            Else
                Console.WriteLine($"指定マクロ '{p}' は存在しません")
            End If
        Next
        Return 0
    End Function
End Class
