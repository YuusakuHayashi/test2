Public Class SaveProjectCommand : Inherits RpaCommandBase

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Call dat.Project.Save(dat.Project.SystemJsonFileName, dat.Project)
        Console.WriteLine($"JsonFile '{dat.Project.SystemJsonFileName}' をセーブしました。")
        Console.WriteLine()
        Return 0
    End Function
End Class
