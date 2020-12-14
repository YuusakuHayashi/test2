﻿Imports System.Runtime.InteropServices
Imports System.IO

Public Class RpaMacroUtility
    Private _MacroFileName As String
    Public Property MacroFileName As String
        Get
            If String.IsNullOrEmpty(Me._MacroFileName) Then
                Me._MacroFileName = $"{RpaProject.SYSTEM_UTILITIES_DIRECTORY}\macro.xlsm"
            End If
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
End Class
