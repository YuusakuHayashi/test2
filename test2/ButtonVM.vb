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
            RaisePropertyChanged("AccessEnableFlag")
        End If
    End Sub

    Private _AccessEnableFlag As Boolean
    Public Property AccessEnableFlag As Boolean
        Get
            Return _AccessEnableFlag
        End Get
        Set(value As Boolean)
            _AccessEnableFlag = value
        End Set
    End Property

    Sub New(ByRef cfvm As ConfigFileVM)
        AddHandler cfvm.PropertyChanged, AddressOf _AccessEnableCheck
    End Sub

    Private Sub AccessCommandExecute(ByVal parameter As Object)

    End Sub
End Class
