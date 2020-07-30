Imports System.ComponentModel
Imports System.Collections.ObjectModel


Public Class AccessButtonVM
    Inherits ViewModel

    Private Sub _AccessEnableCheck(ByVal sender As Object,
                                  ByVal e As PropertyChangedEventArgs)
        Dim cfvm As New ConfigFileVM
        If e.PropertyName <> "ConfigFileName" Then
            Exit Sub
        Else
            cfvm = CType(sender, ConfigFileVM)
            Select Case cfvm.ConfigFileName
                Case cfvm.PLEASE_INPUT
                    AccessEnableFlag = False
                Case vbNullString
                    AccessEnableFlag = False
                Case Else
                    AccessEnableFlag = True
            End Select
        End If
    End Sub


    Private _AccessCommand As ICommand
    Public ReadOnly Property AccessCommand As ICommand
        Get
            If _AccessCommand Is Nothing Then
                _AccessCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _AccessCommandExecute,
                    .CanExecuteHandler = AddressOf _AceessCommandCanExecute
                }
                Return _AccessCommand
            Else
                Return _AccessCommand
            End If
        End Get
    End Property

    Private _AccessEnableFlag As Boolean
    Public Property AccessEnableFlag As Boolean
        Get
            Return _AccessEnableFlag
        End Get
        Set(value As Boolean)
            _AccessEnableFlag = value
            CType(AccessCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _AccessCommandExecute(ByVal parameter As Object)
        MsgBox("hello")
    End Sub

    Private Function _AceessCommandCanExecute(ByVal parameter As Object) As Boolean
        Return AccessEnableFlag
    End Function

    Sub New(ByRef cfvm As ConfigFileVM)
        AddHandler cfvm.PropertyChanged, AddressOf _AccessEnableCheck
    End Sub
End Class
