Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class MenuExplorerViewModel : Inherits ControllerViewModelBase(Of MenuExplorerViewModel)
    Private _MenuIndex As Integer
    Public Property MenuIndex As Integer
        Get
            Return Me._MenuIndex
        End Get
        Set(value As Integer)
            Me._MenuIndex = value
            Call SwitchMainContent()
        End Set
    End Property

    Private _Menus As ObservableCollection(Of ExecuteMenu)
    Public ReadOnly Property Menus As ObservableCollection(Of ExecuteMenu)
        Get
            If Me._Menus Is Nothing Then
                Me._Menus = New ObservableCollection(Of ExecuteMenu)
                Me._Menus.Add(New ExecuteMenu With {.Name = "起動＆チェック", .Content = New RunnerViewModel With {.ViewController = ViewController}})
            End If
            Return Me._Menus
        End Get
    End Property

    Private Sub SwitchMainContent()
        If Me.MenuIndex >= 0 Then
            Dim obj = Me.Menus(Me.MenuIndex).Content
            Call obj.Initialize()
            ViewController.MainContent = obj
        End If
    End Sub

    Sub New()
        Me.MenuIndex = -1
    End Sub
End Class
