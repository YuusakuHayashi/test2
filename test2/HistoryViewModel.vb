Imports System.ComponentModel

Public Class HistoryViewModel
    Inherits ProjectBaseViewModel(Of HistoryViewModel)

    Private _HistoryContents As String
    Public Property HistoryContents As String
        Get
            Return Me._HistoryContents
        End Get
        Set(value As String)
            Me._HistoryContents = value
            RaisePropertyChanged("HistoryContents")
        End Set
    End Property

    Private Sub _HistoryContentsUpdate(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        Dim hist As HistoryModel
        If e.PropertyName = "Contents" Then
            hist = CType(sender, HistoryModel)
            Me.HistoryContents = hist.Contents
        End If
    End Sub

    Private Sub _AddHandlerHistoryChanged()
        AddHandler Model.History.PropertyChanged, AddressOf _HistoryContentsUpdate
    End Sub

    Protected Overrides Sub ContextModelCheck()
        Dim b As Boolean : b = True
        If b Then
            ViewModel.ContextModel = Me
        End If
    End Sub

    Protected Overrides Sub MyInitializing()
        Me.HistoryContents = Model.History.Contents
    End Sub


    Sub New(ByRef m As Model,
            ByRef vm As ViewModel)

        Dim ahp(0) As AddHandlerProxy
        Dim ahp2 As AddHandlerProxy

        ahp(0) = AddressOf Me._AddHandlerHistoryChanged

        ahp2 = [Delegate].Combine(ahp)

        ' ビューモデルの設定
        Initializing(m, vm, ahp2)
    End Sub
End Class
