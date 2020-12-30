Imports System.Reflection
Public Class ChangeProjectPropertyUsingFolderBrowserCommand : Inherits RpaCommandBase
    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        Dim pname As String = trn.Parameters(0)
        Dim pi As PropertyInfo = rpa.GetType().GetProperty(pname)
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

        Dim ptext As String = SetDirectoryFromDialog(trn, rpa, ini, trn.Parameters(0))
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

        pi.SetValue(rpa, pvalue)
        Console.WriteLine($"プロパティ '{pname}' を 値' {ptext}' に変更しました")
        Call rpa.Save(rpa.SystemJsonFileName, rpa)
        Console.WriteLine($"'{rpa.SystemJsonFileName}' を更新しました")
        Console.WriteLine()
    End Function
End Class
