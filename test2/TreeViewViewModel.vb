Imports System.ComponentModel
Imports System.Collections.ObjectModel


Public Class TreeViewViewModel
    Inherits ViewModel

    Private __SqlM As SqlModel
    Private Property _SqlM As SqlModel
        Get
            Return __SqlM
        End Get
        Set(value As SqlModel)
            __SqlM = value
        End Set
    End Property


    'SqlModelの変更を自身に反映
    Private Sub _SqlModelUpdate(ByVal sender As Object,
                                ByVal e As PropertyChangedEventArgs)
        Dim sm As SqlModel
        sm = CType(sender, SqlModel)
        Me._SqlM = sm
        If e.PropertyName = "TreeViewM" Then
            Call _ViewModelUpdate()
        End If
    End Sub


    'SqlModel(TreeViewMプロパティ)の変更を自身に反映
    Private Sub _ViewModelUpdate()
        If Not Me._SqlM.TreeViewM.Equals(Me.Model) Then
            Me.Model = Me._SqlM.TreeViewM
        End If
    End Sub

    'Private Sub _ViewModelUpdate(ByVal sender As Object,
    '                             ByVal e As PropertyChangedEventArgs)
    '    Dim sm As SqlModel

    '    Select Case e.PropertyName
    '        Case "TreeViewM"
    '        Case Else
    '            Exit Sub
    '    End Select


    '    sm = CType(sender, SqlModel)
    '    Select Case e.PropertyName
    '        Case "TreeViewM"
    '            Me._Model = sm.TreeViewM
    '        Case Else
    '            Exit Sub
    '    End Select
    'End Sub


    Private _Model As TreeViewModel
    Public Property Model As TreeViewModel
        Get
            Return _Model
        End Get
        Set(value As TreeViewModel)
            _Model = value
            RaisePropertyChanged("Model")
        End Set
    End Property

    'Private Property _RealName As String
    'Public Property RealName As String
    '    Get
    '        Return _RealName
    '    End Get
    '    Set(value As String)
    '        _RealName = value
    '    End Set
    'End Property

    'Private Property _Checked As Boolean
    'Public Property Checked As Boolean
    '    Get
    '        If Me._Checked = Nothing Then
    '            Return True
    '        Else
    '            Return _Checked
    '        End If
    '    End Get
    '    Set(value As Boolean)
    '        _Checked = value
    '        RaisePropertyChanged("Checked")
    '    End Set
    'End Property


    'Private _Child As List(Of TreeViewViewModel)
    'Public Property Child As List(Of TreeViewViewModel)
    '    Get
    '        Return _Child
    '    End Get
    '    Set(value As List(Of TreeViewViewModel))
    '        _Child = value
    '    End Set
    'End Property
    '-----------------------------------------------------------------------------'


    'Private _ViewModelList As List(Of TreeViewViewModel)
    'Public Property ViewModelList As List(Of TreeViewViewModel)
    '    Get
    '        Return _ViewModelList
    '    End Get
    '    Set(value As List(Of TreeViewViewModel))
    '        _ViewModelList = value
    '    End Set
    'End Property

    'Private _ClickCommand As ICommand
    'Public ReadOnly Property ClickCommand As ICommand
    '    Get
    '        If _ClickCommand Is Nothing Then
    '            _ClickCommand = New DelegateCommand With {
    '                .ExecuteHandler = AddressOf _ClickCommandExecute,
    '                .CanExecuteHandler = AddressOf _ClickCommandCanExecute
    '            }
    '            Return _ClickCommand
    '        Else
    '            Return _ClickCommand
    '        End If
    '    End Get
    'End Property

    'Private _AccessEnableFlag As Boolean
    'Public Property AccessEnableFlag As Boolean
    '    Get
    '        Return _AccessEnableFlag
    '    End Get
    '    Set(value As Boolean)
    '        _AccessEnableFlag = value
    '        'RaisePropertyChanged("AccessEnableFlag")
    '        CType(ClickCommand, DelegateCommand).RaiseCanExecuteChanged()
    '    End Set
    'End Property

    'Private Sub _ClickCommandExecute(ByVal parameter As Object)
    '    MsgBox("hello")
    'End Sub

    'Private Function _ClickCommandCanExecute(ByVal parameter As Object) As Boolean
    '    Return True
    'End Function


    'Sub New()
    Sub New(ByRef sm As SqlModel)
        Me._SqlM = sm
        Me.Model = sm.TreeViewM

        AddHandler Me._SqlM.PropertyChanged, AddressOf _SqlModelUpdate
        'AddHandler Me._SqlM.TreeViewM.PropertyChanged, AddressOf _ViewModelUpdate

        'Dim a As String
        'Dim b As String
        'Dim c As String

        'If tvml Is Nothing Then
        '    Me._ModelList = Nothing
        'Else
        '    For Each tvm In tvml
        '        Me.RealName = tvm.RealName
        '        Me.Checked = tvm.Checked
        '        Me.Child = tvm.Child
        '        Me.ModelList.Add(Me)
        '    Next
        'End If
        'AddHandler Me.PropertyChanged, AddressOf Me._ModelUpdate

        'Me.ModelList = New List(Of TreeViewModel) From {
        '    New TreeViewModel With {
        '        .RealName = "AAA", .Child = New List(Of TreeViewModel) From {
        '            New TreeViewModel With {
        '                .RealName = "ABA"
        '            },
        '            New TreeViewModel With {
        '                .RealName = "ACA"
        '            }
        '        }
        '    },
        '    New TreeViewModel With {
        '        .RealName = "BAA", .Child = New List(Of TreeViewModel) From {
        '            New TreeViewModel With {
        '                .RealName = "BBA"
        '            },
        '            New TreeViewModel With {
        '                .RealName = "BCA"
        '            }
        '        }
        '    }
        '}
        'm.ModelSave()
        'Dim s As ServerModel
        'Me._VM = New Model
        's = Me._VM.ServerLoad
        'Me.ServerVM = New List(Of ServerModel)
        'Me.ServerVM.Add(s)
    End Sub
End Class
