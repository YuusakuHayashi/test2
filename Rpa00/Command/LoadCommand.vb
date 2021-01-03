Imports System.IO

Public Class LoadCommand : Inherits RpaCommandBase

    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides Function CanExecute(ByRef dat As RpaDataWrapper) As Boolean
        If dat.Project.Count Is Nothing Then
            Console.WriteLine($"プロジェクトがありません")
            Console.WriteLine()
            Return False
        End If

        If dat.Initializer.Projects.Count = 0 Then
            Console.WriteLine($"プロジェクトの登録がありません")
            Console.WriteLine()
            Return False
        End If
        Return True
    End Function

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
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
            Console.WriteLine($"ロードするプロジェクトを選択してください")
            idxtext = dat.Transaction.ShowRpaIndicator(dat)
            If IsNumeric(idxtext) Then
                idx2 = Integer.Parse(idxtext)
            Else
                idx2 = lastidx + 1
            End If
            Console.WriteLine()
        Loop Until idx2 <= lastidx

        If idx2 = lastidx Then
            Return 0
        End If

        Dim [load] As Object = RpaModule.RpaObject(dat.Initializer.Projects(idx2).Architecture)
        [load] = [load].Load(dat.Initializer.Projects(idx2).JsonFileName)

        If [load] Is Nothing Then
            Console.WriteLine($"JsonFile '{dat.Project.SystemJsonFileName}' のロードに失敗しました。")
            Console.WriteLine()
            Return 1000
        End If

        dat.Initializer.CurrentProject = dat.Initializer.Projects(idx2)
        dat.Project = [load]
        Console.WriteLine($"JsonFile '{dat.Project.SystemJsonFileName}' をロードしました。")
        Console.WriteLine()
        Return 0
    End Function
End Class
