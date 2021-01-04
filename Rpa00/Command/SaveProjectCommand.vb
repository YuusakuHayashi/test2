Public Class SaveProjectCommand : Inherits RpaCommandBase

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        'Call dat.Project.Save(dat.Project.SystemJsonFileName, dat.Project)
        Call RpaModule.Save(dat.Project.SystemJsonFileName, dat.Project, dat.Project.SystemJsonChangeFileName)
        Console.WriteLine($"JsonFile '{dat.Project.SystemJsonFileName}' をセーブしました。")
        Console.WriteLine()
        Return 0
    End Function
End Class
