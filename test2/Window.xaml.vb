Imports System.Collections.ObjectModel

Public Class Window
    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Dim ml As New JsonHandler(Of Object)

        Dim app As AppDirectoryModel
        Dim m As New Model
        Dim vm As New ViewModel
        Dim project As New ProjectInfoModel

        app = ml.ModelLoad(Of AppDirectoryModel)(AppDirectoryModel.ModelFileName)
        If app Is Nothing Then
            app = New AppDirectoryModel
        End If
        app.ModelSave(AppDirectoryModel.ModelFileName, app)

        Dim udvm As New UserDirectoryViewModel
        Call udvm.Initialize(m, vm, app, project)

        Me.MainFlame.DataContext = vm
    End Sub
End Class
