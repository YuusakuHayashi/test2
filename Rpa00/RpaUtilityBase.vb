Public MustInherit Class RpaUtilityBase
    Public MustOverride ReadOnly Property ExecuteHandler(trn As RpaTransaction, rpa As RpaProject) As RpaSystem.ExecuteDelegater
End Class
