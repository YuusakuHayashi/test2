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

    '---------------------------------------------------------------------------------------------'
    Private ReadOnly Property CheckAndRunIconFileName As String
        Get
            Return $"{Controller.SystemDllDirectory}\checkandrun.png"
        End Get
    End Property
    Private _CheckAndRunIcon As BitmapImage
    Private ReadOnly Property CheckAndRunIcon As BitmapImage
        Get
            If Me._CheckAndRunIcon Is Nothing Then
                Me._CheckAndRunIcon = RpaGuiModule.CreateIcon(Me.CheckAndRunIconFileName)
            End If
            Return Me._CheckAndRunIcon
        End Get
    End Property
    '---------------------------------------------------------------------------------------------'

    Private _Menus As ObservableCollection(Of ExecuteMenu)
    Public ReadOnly Property Menus As ObservableCollection(Of ExecuteMenu)
        Get
            If Me._Menus Is Nothing Then
                Dim [obj] As Object
                Dim [men] As ExecuteMenu
                Me._Menus = New ObservableCollection(Of ExecuteMenu)

                '---------------------------------------------------------------------------------'
                [obj] = New RunnerViewModel With {.ViewController = ViewController}
                [obj].Initialize()
                [men] = New ExecuteMenu With {.Icon = Me.CheckAndRunIcon, .Name = "起動とチェック", .Content = [obj]}
                Me._Menus.Add([men])
                '---------------------------------------------------------------------------------'
            End If
            Return Me._Menus
        End Get
    End Property

    Private Sub SwitchMainContent()
        If Me.MenuIndex >= 0 Then
            Dim obj = Me.Menus(Me.MenuIndex).Content
            'Call obj.Initialize()
            ViewController.MainContent = obj
        End If
    End Sub

    Sub New()
        Me.MenuIndex = -1
    End Sub
End Class
