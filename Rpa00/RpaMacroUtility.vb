Imports System.Runtime.InteropServices
Imports System.IO

Public Class RpaMacroUtility
    
    Public ReadOnly Property ExecuteHandler(trn As RpaTransaction, rpa As RpaProject) As RpaSystem.ExecuteDelegater
        Get
            Dim dlg As RpaSystem.ExecuteDelegater
            Select Case trn.MainCommand
                Case "InstallMacro" : dlg = AddressOf UpdateMacro
                Case "UpdateMacro" : dlg = AddressOf UpdateMacro
                Case Else : dlg = Nothing
            End Select
            Return dlg
        End Get
    End Property

    Private ReadOnly Property _MacroUtilityDirectory As String
        Get
            Dim d = $"{RpaProject.SYSTEM_DIRECTORY}\MacroUtility"
            If Not Directory.Exists(d) Then
                Directory.CreateDirectory(d)
            End If
            Return d
        End Get
    End Property

    Private ReadOnly Property _MacroFileName As String
        Get
            Return $"{RpaProject.SYSTEM_DIRECTORY}\macro.xlsm"
        End Get
    End Property

    Public Sub InvokeMacro(ByVal method As String, args() As Object)
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

        If Not File.Exists(Me._MacroFileName) Then
            Console.WriteLine($"マクロファイル '{Me._MacroFileName}' が存在しません")
            Console.WriteLine($"開発者から入手してください")
            Return 1000
        End If

        For Each p In trn.Parameters
            bas = $"{Me._MacroUtilityDirectory}\{p}"
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
