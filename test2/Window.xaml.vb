Imports System.Collections.ObjectModel

Public Class Window
    Sub New()
        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Dim ml As New JsonHandler(Of Object)
        Dim vm As New ViewModel

        Dim app As New AppDirectoryModel
        app = ml.ModelLoad(Of AppDirectoryModel)(AppDirectoryModel.ModelFileName)
        If app Is Nothing Then
            app = New AppDirectoryModel
        End If
        app.ModelSave(AppDirectoryModel.ModelFileName, app)
        app.Initialize()

        Dim vc As New ViewController
        Call vc.Initialize(app, vm)

        Dim udvm As New UserDirectoryViewModel
        Call udvm.Initialize(app, vm)

        Me.DataContext = vm
    End Sub
End Class
