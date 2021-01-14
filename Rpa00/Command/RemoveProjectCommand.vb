Imports System.IO
Public Class RemoveProjectCommand : Inherits RpaCommandBase
    Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        If dat.Initializer.Projects.Count = 0 Then
            Console.WriteLine($"プロジェクトの登録がありません")
            Console.WriteLine()
            Return False
        End If
        Return True
    End Function

    Private Function Main(ByRef dat As RpaDataWrapper) As Integer
        Dim idx As Integer = -1
        Dim lastidx As Integer = -1
        For Each prj As RpaInitializer.RpaProject In dat.Initializer.Projects
            idx = dat.Initializer.Projects.IndexOf(prj)
            Console.WriteLine($"{idx} {prj.Name} {prj.ProjectDirectory}")
        Next
        lastidx = idx + 1
        Console.WriteLine($"{lastidx} やっぱりやめる")
        Console.WriteLine()

        Dim idx2 As Integer = -1
        Dim idxtext As String = vbNullString
        Do
            Console.WriteLine($"削除するプロジェクトを選択してください")
            idxtext = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
            idx2 = IIf(IsNumeric(idxtext), Integer.Parse(idxtext), (lastidx + 1))
        Loop Until idx2 <= lastidx
        If idx2 = lastidx Then
            Console.WriteLine()
            Return 1000
        End If

        Dim target As RpaInitializer.RpaProject = dat.Initializer.Projects(idx)
        Dim pname As String = target.Name
        Dim pdir As String = target.ProjectDirectory

        Dim yorn As String = vbNullString
        Do
            yorn = vbNullString
            Console.WriteLine($"以下のプロジェクトを削除していよろしいですか？ (y/n)")
            Console.WriteLine($"プロジェクト名 '{pname}'")
            Console.WriteLine($"場所           '{pdir}'")
            yorn = dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
        Loop Until yorn = "y" Or yorn = "n"

        If yorn = "n" Then
            Return 0
        End If

        Directory.Delete(target.ProjectDirectory, True)
        Console.WriteLine($"プロジェクト '{pname}' を削除しました")

        If dat.Initializer.CurrentProject.Name = target.Name Then
            dat.Initializer.CurrentProject = Nothing
        End If
        dat.Initializer.Projects.Remove(target)

        dat.Initializer.Save(RpaCui.SystemIniFileName, dat.Initializer)
        Console.WriteLine($"ファイル '{RpaCui.SystemIniFileName}' をセーブしました")
        Console.WriteLine()
        Return 0
    End Function

    Sub New()
        Me.ExecuteHandler = AddressOf Main
    End Sub
End Class
