Public Class SaveInitializerCommand : Inherits RpaCommandBase
    Public Overrides ReadOnly Property ExecuteIfNoProject As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Call dat.Project.Save(RpaInitializer.SystemIniFileName, dat.Initializer)
        Console.WriteLine($"IniFile '{RpaInitializer.SystemIniFileName}' をセーブしました。")
        Console.WriteLine()
        Return 0
    End Function
End Class
