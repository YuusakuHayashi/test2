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

    'Private Sub GridSplitter_Drop(sender As Object, e As DragEventArgs)
    '    DelegateEventListener.Instance.RaiseMultiViewRowGridSplitterChanged(sender)
    'End Sub

    'Private Sub GridSplitter_DragOver(sender As Object, e As DragEventArgs)        Me.r
    '    DelegateEventListener.Instance.RaiseMultiViewRowGridSplitterChanged(sender)
    'End Sub

    Private Sub ExplorerView_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        DelegateEventListener.Instance.RaiseMultiViewSizeChanged(sender)
    End Sub

    Private Sub MainView_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        DelegateEventListener.Instance.RaiseMultiViewSizeChanged(sender)
    End Sub

    Private Sub HistoryView_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        DelegateEventListener.Instance.RaiseMultiViewSizeChanged(sender)
    End Sub

    'Private Sub GridSplitter_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.Key = Key.F2 Then
    '        DelegateEventListener.Instance.RaiseMultiViewRowGridSplitterChanged(sender)
    '    End If
    'End Sub

    'Private Sub GridSplitter_KeyUp(sender As Object, e As KeyEventArgs)
    '    If e.Key = Key.F2 Then
    '        DelegateEventListener.Instance.RaiseMultiViewRowGridSplitterChanged(sender)
    '    End If
    'End Sub
End Class
