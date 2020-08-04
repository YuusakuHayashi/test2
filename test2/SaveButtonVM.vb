Imports System.ComponentModel

Public Class SaveButtonVM
    Inherits ViewModel

    Private __Model As SqlModel
    Private Property _Model As SqlModel
        Get
            Return __Model
        End Get
        Set(value As SqlModel)
            __Model = value
        End Set
    End Property


    Private __VM As SqlStatusVM
    Private Property _VM As SqlStatusVM
        Get
            Return __VM
        End Get
        Set(value As SqlStatusVM)
            __VM = value
        End Set
    End Property


    Private __ABVM As AccessButtonVM
    Private Property _ABVM As AccessButtonVM
        Get
            Return __ABVM
        End Get
        Set(value As AccessButtonVM)
            __ABVM = value
        End Set
    End Property


    'SqlModel変更の反映, プロパティの意識無し
    Private Sub _ModelUpdate(ByVal sender As Object,
                              ByVal e As PropertyChangedEventArgs)
        Dim sm As SqlModel

        sm = CType(sender, SqlModel)
        Me._Model = sm
    End Sub


    'SqlStatusVM変更の反映, プロパティの意識無し
    Private Sub _VMUpdate(ByVal sender As Object,
                              ByVal e As PropertyChangedEventArgs)
        Dim ssvm As SqlStatusVM

        ssvm = CType(sender, SqlStatusVM)
        Me._VM = ssvm
    End Sub


    'AccessButtonVM変更の反映, プロパティの意識無し
    Private Sub _ABVMUpdate(ByVal sender As Object,
                            ByVal e As PropertyChangedEventArgs)
        Dim abvm As AccessButtonVM

        abvm = CType(sender, AccessButtonVM)
        Me._ABVM = abvm
    End Sub



    Private Sub _SaveEnableCheck(ByVal sender As Object,
                                  ByVal e As PropertyChangedEventArgs)
        Dim abvm As AccessButtonVM

        Select Case e.PropertyName
            Case "AccessEnableFlag"
            Case Else
                Exit Sub
        End Select

        abvm = CType(sender, AccessButtonVM)
        Me.SaveEnableFlag = abvm.AccessEnableFlag
    End Sub



    Private _SaveCommand As ICommand
    Public ReadOnly Property SaveCommand As ICommand
        Get
            If _SaveCommand Is Nothing Then
                _SaveCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _SaveCommandExecute,
                    .CanExecuteHandler = AddressOf _SaveCommandCanExecute
                }
                Return _SaveCommand
            Else
                Return _SaveCommand
            End If
        End Get
    End Property

    Private _SaveEnableFlag As Boolean
    Public Property SaveEnableFlag As Boolean
        Get
            Return _SaveEnableFlag
        End Get
        Set(value As Boolean)
            _SaveEnableFlag = value
            CType(SaveCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _SaveCommandExecute(ByVal parameter As Object)
        Call Me._Model.Save()
    End Sub

    Private Function _SaveCommandCanExecute(ByVal parameter As Object) As Boolean
        Return SaveEnableFlag
    End Function


    Sub New(ByRef sm As SqlModel, ByRef ssvm As SqlStatusVM, ByRef abvm As AccessButtonVM)

        Me._Model = sm
        Me._VM = ssvm
        Me._ABVM = abvm

        AddHandler sm.PropertyChanged, AddressOf Me._ModelUpdate
        AddHandler ssvm.PropertyChanged, AddressOf Me._VMUpdate
        AddHandler abvm.PropertyChanged, AddressOf Me._ABVMUpdate
        AddHandler abvm.PropertyChanged, AddressOf Me._SaveEnableCheck
    End Sub
End Class
