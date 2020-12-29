Imports System.Reflection

Public Class SetupInitializerCommand : Inherits RpaCommandBase
    Private ReadOnly Property ExecuteHandler(trn As RpaTransaction, rpa As Object, ini As RpaInitializer) As Object
        Get
            Dim cmd As Object
            Select Case trn.MainCommand
                Case "change" : cmd = New ChangeInitializerCommand
                ' 現状、SetupInitializerのExit と ApplicationのExit に差はないため再利用している
                Case "exit" : cmd = New ExitCommand
                Case Else : cmd = Nothing
            End Select
            Return cmd
        End Get
    End Property
    '---------------------------------------------------------------------------------------------'

    Public Overrides Function Execute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
        trn.Modes.Add("initializer")

        ' 初回時にプロパティ一覧を表示
        Dim props As PropertyInfo() = ini.GetType().GetProperties()
        For Each prop In props
            Console.WriteLine($"{prop.Name} {prop.PropertyType.Name} {prop.GetValue(ini).ToString}")
        Next

        Do
            trn.CommandText = trn.ShowRpaIndicator(rpa)
            Call trn.CreateCommand()
            Call CommandExecute(trn, rpa, ini)
            Console.WriteLine()
        Loop Until trn.ExitFlag

        trn.ExitFlag = False
        trn.Modes.Remove(trn.Modes.Remove("initializer"))
    End Function

    Public Sub CommandExecute(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) 
        Dim cmd = ExecuteHandler(trn, rpa, ini)
        If cmd IsNot Nothing Then
            trn.ReturnCode = cmd.Execute(trn, rpa, ini)
        Else
            Console.WriteLine($"コマンド : '{trn.CommandText}' はありません")
        End If
    End Sub

    'Private Function ChangeInitializer(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer) As Integer
    '    Dim txt As String = trn.ShowRpaIndicator(rpa)
    '    Dim param As String() = txt.Split(" "c)

    '    Dim pname As String = param(0)
    '    Dim pi As PropertyInfo = ini.GetType().GetProperty(pname)
    '    If pi Is Nothing Then
    '        Console.WriteLine($"プロパティ '{pname}' は存在しません")
    '        Exit Function
    '    End If

    '    Dim ptext As String = param(1)
    '    Dim pvalue As Object
    '    Select Case pi.PropertyType
    '        Case (New Integer).GetType             ' 数値
    '            pvalue = Integer.Parse(ptext)
    '        Case "String".GetType                  ' 文字列
    '            pvalue = ptext
    '        Case True.GetType                      ' 論理値
    '            pvalue = Boolean.Parse(ptext)
    '        Case Else
    '            Console.WriteLine($"プロパティ '{pname}' は変更不可です")
    '            pvalue = Nothing
    '    End Select

    '    If pvalue Is Nothing Then
    '        Console.WriteLine($"プロパティ '{pname}' を 値' {ptext}' に変更できませんでした")
    '        Exit Function
    '    End If

    '    pi.SetValue(ini, pvalue)
    '    Console.WriteLine($"プロパティ '{pname}' を 値' {ptext}' に変更しました")
    '    Call ini.Save(CommonProject.SystemIniFileName, ini)
    '    Console.WriteLine($"'{CommonProject.SystemIniFileName}' を更新しました")
    'End Function
End Class
