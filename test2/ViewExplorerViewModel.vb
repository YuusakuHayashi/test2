﻿Imports System.Collections.ObjectModel
Imports test2
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ViewExplorerViewModel
    Inherits BaseViewModel2

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        Call BaseInitialize(app, vm)
    End Sub
End Class
