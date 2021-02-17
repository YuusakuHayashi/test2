Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Dynamic

Public Class ConsoleOutputViewModel : Inherits ControllerViewModelBase(Of ConsoleOutputViewModel)
    Private _GuiCommandTexts As ObservableCollection(Of String)
    Public Property GuiCommandTexts As ObservableCollection(Of String)
        Get
            If Me._GuiCommandTexts Is Nothing Then
                Me._GuiCommandTexts = New ObservableCollection(Of String)
            End If
            Return Me._GuiCommandTexts
        End Get
        Set(value As ObservableCollection(Of String))
            Me._GuiCommandTexts = value
        End Set
    End Property

    Private _GuiCommandText As String
    Public Property GuiCommandText As String
        Get
            Return Me._GuiCommandText
        End Get
        Set(value As String)
            Me._GuiCommandText = value
            RaisePropertyChanged("GuiCommandText")
            If ViewController.GuiCommandTextPath <> value Then
                ViewController.GuiCommandTextPath = value
            End If
        End Set
    End Property

    '---------------------------------------------------------------------------------------------'
    Private _ExecuteGuiCommandCommand As ICommand
    Public ReadOnly Property ExecuteGuiCommandCommand As ICommand
        Get
            If Me._ExecuteGuiCommandCommand Is Nothing Then
                Me._ExecuteGuiCommandCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf ExecuteGuiCommandCommandExecute,
                    .CanExecuteHandler = AddressOf ExecuteGuiCommandCommandCanExecute
                }
            End If
            Return Me._ExecuteGuiCommandCommand
        End Get
    End Property

    Private _IsExecuteGuiCommandCommandEnabled As Boolean
    Public Property IsExecuteGuiCommandCommandEnabled As Boolean
        Get
            Return Me._IsExecuteGuiCommandCommandEnabled
        End Get
        Set(value As Boolean)
            Me._IsExecuteGuiCommandCommandEnabled = value
            RaisePropertyChanged("IsExecuteGuiCommandCommandEnabled")
            CType(ExecuteGuiCommandCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Async Sub ExecuteGuiCommandCommandExecute(ByVal parameter As Object)
        Me.IsExecuteGuiCommandCommandEnabled = False

        Dim t As Task = Task.Run(
            Sub()
                Call _ExecuteGuiCommandCommandExecute(parameter)
            End Sub
        )
        Dim t2 As Task = Task.Run(
            Sub()
                ViewController.GuiCommandExecuteStatusSender = 1
            End Sub
        )
        Await Task.WhenAll(t, t2)
        Me.IsExecuteGuiCommandCommandEnabled = True

        ' Controllerへの通知
        ViewController.GuiCommandExecuteStatusSender = 0
        Dim i As Integer = AddGuiCommandText()

        ' Controllerへの通知
        'ViewController.CommandLogsPath = ViewController.Data.System.CommandLogs
    End Sub

    Private Sub _ExecuteGuiCommandCommandExecute(ByVal parameter As Object)
        ViewController.Data.System.ExecuteMode = Rpa00.RpaSystem.ExecuteModeNumber.RpaGui
        ViewController.Data.System.GuiCommandText = $"{Me.GuiCommandText}"
        Call ViewController.Data.System.Main(ViewController.Data)
    End Sub

    Private Function ExecuteGuiCommandCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me.IsExecuteGuiCommandCommandEnabled
    End Function

    Private Function CheckExecuteGuiCommandCommandEnabled() As Integer
        If String.IsNullOrEmpty(Me.GuiCommandText) Then
            Return False
        End If

        Dim cmds As String() = Me.GuiCommandText.Split(" "c)
        If Not ViewController.ExecutableGuiCommands.Contains(cmds(0)) Then
            Return False
        End If
        Return True
    End Function
    '---------------------------------------------------------------------------------------------'

    Private Function AddGuiCommandText() As Integer
        If Not Me.GuiCommandTexts.Contains(Me.GuiCommandText) Then
            Me.GuiCommandTexts.Add(Me.GuiCommandText)
        End If
        If Not ViewController.GuiCommandTextsPath.Contains(Me.GuiCommandText) Then
            ViewController.GuiCommandTextsPath.Add(Me.GuiCommandText)
        End If
        Return 0
    End Function

    Private Sub CheckPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If e.PropertyName = "GuiCommandText" Then
            Me.IsExecuteGuiCommandCommandEnabled = CheckExecuteGuiCommandCommandEnabled()
            Exit Sub
        End If
    End Sub

    Private Sub SystemPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If e.PropertyName = "StatusCode" Then
            ViewController.StatusCode = ViewController.Data.System.StatusCode
            Exit Sub
        End If
    End Sub

    Private Sub CheckViewControllerPropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        If e.PropertyName = "GuiCommandTextPath" Then
            Me.GuiCommandText = ViewController.GuiCommandTextPath
            Exit Sub
        End If
        If e.PropertyName = "ExecuteGuiCommandPath" Then
            If ViewController.ExecuteGuiCommandPath Then
                Call ExecuteGuiCommandCommandExecute(Nothing)
                ViewController.ExecuteGuiCommandPath = False
            End If
            Exit Sub
        End If
    End Sub

    Public Overrides Sub Initialize()
        AddHandler ViewController.PropertyChanged, AddressOf CheckViewControllerPropertyChanged
        AddHandler ViewController.Data.System.PropertyChanged, AddressOf SystemPropertyChanged
    End Sub

    Sub New()
        AddHandler Me.PropertyChanged, AddressOf CheckPropertyChanged
    End Sub
End Class
