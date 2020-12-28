Imports System.Reflection

Public Class SetupInitializerCommand : Inherits RpaCommandBase
    Private Delegate Function ExecuteDelegater(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
    Private ReadOnly Property ExecuteHandler(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As ExecuteDelegater
        Get
            Dim cmd As ExecuteDelegater
            Select Case trn.MainCommand
                Case "change" : cmd = AddressOf ChangeInitializer
                Case "exit"
                Case Else : cmd = Nothing
            End Select
            Return cmd
        End Get
    End Property
    '---------------------------------------------------------------------------------------------'

    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        Dim cmd As Object
        Dim i As Integer
        trn.Modes.Add("setup")
        Do
            trn.CommandText = trn.ShowRpaIndicator(rpa)
            cmd = ExecuteHandler(trn, rpa, ini)
            i = cmd(trn, rpa, ini)
            Console.WriteLine()
        Loop Until trn.ExitFlag

        trn.ExitFlag = False
        trn.Modes.Remove(trn.Modes.Remove("setup"))
    End Function

    Private Function ChangeInitializer(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        Dim tp As Type = ini.GetType()
        Dim props As PropertyInfo() = tp.GetProperties()
        For Each prop In props
            Console.WriteLine($"{prop.Name} {prop.PropertyType.Name} {prop.GetValue(ini).ToString}")
        Next

        Dim txt As String = trn.ShowRpaIndicator(rpa)
        Dim param As String() = txt.Split(" "c)

        Dim pname As String = param(0)
        Dim pi As PropertyInfo = ini.GetType().GetProperty(pname)
        If pi Is Nothing Then
            Console.WriteLine($"プロパティ '{pname}' は存在しません")
            Exit Function
        End If

        Dim ptext As String = param(1)
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
            Exit Function
        End If

        pi.SetValue(ini, pvalue)
        Console.WriteLine($"プロパティ '{pname}' を 値' {ptext}' に変更しました")
        Call ini.Save(CommonProject.SystemIniFileName, ini)
        Console.WriteLine($"'{CommonProject.SystemIniFileName}' を更新しました")
    End Function
End Class
