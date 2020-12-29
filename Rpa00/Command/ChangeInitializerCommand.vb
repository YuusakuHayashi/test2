Imports Rpa00
Imports System.Reflection

Public Class ChangeInitializerCommand : Inherits RpaCommandBase
    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        Dim pname As String = trn.Parameters(0)
        Dim pi As PropertyInfo = ini.GetType().GetProperty(pname)
        If pi Is Nothing Then
            Console.WriteLine($"プロパティ '{pname}' は存在しません")
            Return 1000
        End If

        Dim ptext As String = trn.Parameters(1)
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
                pvalue = Nothing
        End Select

        If pvalue Is Nothing Then
            Console.WriteLine($"プロパティ '{pname}' を 値' {ptext}' に変更できませんでした")
            Return 1000
        End If

        pi.SetValue(ini, pvalue)
        Console.WriteLine($"プロパティ '{pname}' を 値' {ptext}' に変更しました")
        Call ini.Save(CommonProject.SystemIniFileName, ini)
        Console.WriteLine($"'{CommonProject.SystemIniFileName}' を更新しました")
    End Function
End Class
