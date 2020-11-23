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

        Dim vc As New ViewController
        Call vc.Initialize(app, vm)

        Dim udvm As New UserDirectoryViewModel
        Call udvm.Initialize(app, vm)

        ' DynamicView に統一予定
        Me.DataContext = vm
    End Sub

    'Private Sub ContentPanel_SizeChanged(sender As Object, e As SizeChangedEventArgs)
    '    DelegateEventListener.Instance.RaiseViewResized(sender, e)
    'End Sub

    'Private Sub RightPanel_SizeChanged(sender As Object, e As SizeChangedEventArgs)
    '    DelegateEventListener.Instance.RaiseViewResized(sender, e)
    'End Sub

    'Private Sub BottomPanel_SizeChanged(sender As Object, e As SizeChangedEventArgs)
    '    DelegateEventListener.Instance.RaiseViewResized(sender, e)
    'End Sub

    Private Sub MenuContent_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        DelegateEventListener.Instance.RaiseViewResized(sender, e)
    End Sub

    Private Sub LeftExplorerContent_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        DelegateEventListener.Instance.RaiseViewResized(sender, e)
    End Sub

    Private Sub MainContent_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        DelegateEventListener.Instance.RaiseViewResized(sender, e)
    End Sub

    Private Sub RightExplorerContent_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        DelegateEventListener.Instance.RaiseViewResized(sender, e)
    End Sub

    Private Sub HistoryContent_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        DelegateEventListener.Instance.RaiseViewResized(sender, e)
    End Sub
End Class
