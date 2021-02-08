Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Public Class ControllerViewModel : Inherits ViewModelBase

    ' !!! RpaDataWrapper型のdllオブジェクトがキャストされる
    ' 別プロジェクトのせいか、RpaDataWrapper型としては定義できなかった
    ' (するとキャスト時エラーになる)
    Private _Data As Object
    Public Property Data As Object
        Get
            Return Me._Data
        End Get
        Set(value As Object)
            Me._Data = value
        End Set
    End Property

    Private _Content As Object
    Public Property Content As Object
        Get
            Return Me._Content
        End Get
        Set(value As Object)
            Me._Content = value
            Call RaisePropertyChanged("Content")
        End Set
    End Property

    Private _Output As Rpa00.OutputData
    Public Property Output As Rpa00.OutputData
        Get
            Return Me._Output
        End Get
        Set(value As Rpa00.OutputData)
            Me._Output = value
            RaisePropertyChanged("Output")
        End Set
    End Property

    Private _SelectedMenu As ObservableCollection(Of ExecuteMenu)
    Public ReadOnly Property SelectedMenu As ObservableCollection(Of ExecuteMenu)
        Get
            If Me._SelectedMenu Is Nothing Then
                Dim i As Integer = -1
                Me._SelectedMenu = New ObservableCollection(Of ExecuteMenu)
                Me._SelectedMenu.Add(New ExecuteMenu With {.Name = "起動＆チェック", .DataObject = (New RunnerViewModel(Data))}) : AddHandler Me._SelectedMenu((i + 1)).PropertyChanged, AddressOf CheckSwitchMenu
            End If
            Return Me._SelectedMenu
        End Get
    End Property

    Private Overloads Sub CheckSwitchMenu(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If sender.GetType.Name = "ExecuteMenu" Then
            If e.PropertyName = "IsSelected" Then
                If sender.IsSelected Then
                    Dim menu As ExecuteMenu = CType(sender, ExecuteMenu)
                    Dim i As Integer = SwitchMenu(menu)
                End If
            End If
        End If
    End Sub

    Private Function SwitchMenu(ByRef menu As ExecuteMenu) As Integer
        Me.Content = menu.DataObject
        Return 0
    End Function

    Sub New()
    End Sub
End Class
