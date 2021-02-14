﻿Imports System.Reflection
Public Class ChangeProjectPropertyUsingFolderBrowserCommand : Inherits RpaCommandBase
    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim pname As String = dat.System.CommandData.Parameters(0)
        Dim pi As PropertyInfo = dat.Project.GetType().GetProperty(pname)
        If pi Is Nothing Then
            Console.WriteLine($"プロパティ '{pname}' は存在しません")
            Console.WriteLine()
            Return 1000
        End If

        If Not pi.CanWrite Then
            Console.WriteLine($"プロパティ '{pname}' は変更不可です")
            Console.WriteLine()
            Return 1000
        End If

        Dim ptext As String = SetDirectoryFromDialog(dat, pname)
        If String.IsNullOrEmpty(ptext) Then
            Console.WriteLine($"選択パスが不正です")
            Console.WriteLine()
            Return 1000
        End If

        Dim pvalue As Object
        Select Case pi.PropertyType
            Case (New Integer).GetType             ' 数値
                pvalue = Integer.Parse(ptext)
            Case "String".GetType                  ' 文字列
                pvalue = ptext
            Case True.GetType                      ' 論理値
                pvalue = Boolean.Parse(ptext)
            Case Else
                Console.WriteLine($"プロパティ '{pname}' は変更不可です")
                Console.WriteLine()
                pvalue = Nothing
        End Select

        If pvalue Is Nothing Then
            Console.WriteLine($"プロパティ '{pname}' を 値' {ptext}' に変更できませんでした")
            Console.WriteLine()
            Return 1000
        End If

        pi.SetValue(dat.Project, pvalue)
        Console.WriteLine($"プロパティ '{pname}' を 値' {ptext}' に変更しました")
        Console.WriteLine()
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
        Me.ExecutableParameterCount = {1, 1}
    End Sub
End Class
