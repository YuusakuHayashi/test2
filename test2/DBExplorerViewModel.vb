Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized

Public Class DBExplorerViewModel : Inherits ViewModel

    Private Const _DBEXPLORERVIEW_FILE As String = "DBExplorerView.txt"

    Private _Server As ServerModel
    Public Property Server As ServerModel
        Get
            Return Me._Server
        End Get
        Set(value As ServerModel)
            Me._Server = value
            RaisePropertyChanged("Server")
        End Set
    End Property



    Private _ClickCommand As ICommand
    Public ReadOnly Property ClickCommand As ICommand
        Get
            If Me._ClickCommand Is Nothing Then
                Me._ClickCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _ClickCommandExecute,
                    .CanExecuteHandler = AddressOf _ClickCommandCanExecute
                }
                Return Me._ClickCommand
            Else
                Return Me._ClickCommand
            End If
        End Get
    End Property

    Private Sub _CheckClickEnableFlag()
        Me._ClickEnableFlag = True
    End Sub


    'コマンド実行可否のフラグ
    Private __ClickEnableFlag As Boolean
    Public Property _ClickEnableFlag As Boolean
        Get
            Return Me.__ClickEnableFlag
        End Get
        Set(value As Boolean)
            Me.__ClickEnableFlag = value
            RaisePropertyChanged("ClickEnableFlag")
            CType(ClickCommand, DelegateCommand).RaiseCanExecuteChanged()
        End Set
    End Property

    Private Sub _ClickCommandExecute(ByVal parameter As Object)
    End Sub

    Private Function _ClickCommandCanExecute(ByVal parameter As Object) As Boolean
        Return Me._ClickEnableFlag
    End Function

    Private _IsSelected As Boolean
    Public Property IsSelected As Boolean
        Get
            Return Me._IsSelected
        End Get
        Set(value As Boolean)
            Me._IsSelected = value
        End Set
    End Property


    Sub _SetTest()
        '外部に記録
        'Dim pm As New ProjectModel
        'Dim proxy As ProjectCheckProxy
        'proxy = AddressOf pm.FileCheck
        'Select Case pm.ProjectCheck(proxy, _DBEXPLORERVIEW_FILE)
        '    Case 0
        '        pm.ModelSave(Of ServerModel)(_DBEXPLORERVIEW_FILE, Me.Server)
        '    Case 1
        '        pm.FileEstablish(_DBEXPLORERVIEW_FILE)
        '        pm.ModelSave(Of ServerModel)(_DBEXPLORERVIEW_FILE, Me.Server)
        '    Case 99
        '    Case 999
        'End Select

        'Dim test As ServerModel
        'test = New ServerModel With {
        '    .Name = "root",
        '    .DataBases = New List(Of DataBaseModel) From {
        '        New DataBaseModel With {
        '            .Name = "hoge",
        '            .DataTables = New List(Of DataTableModel) From {
        '                New DataTableModel With {.Name = "hogehoge"},
        '                New DataTableModel With {.Name = "hogefuga"},
        '                New DataTableModel With {.Name = "hogepiyo"}
        '            }
        '        },
        '        New DataBaseModel With {
        '            .Name = "fuga",
        '            .DataTables = New List(Of DataTableModel) From {
        '                New DataTableModel With {.Name = "fugafuga"}
        '            }
        '        },
        '        New DataBaseModel With {
        '            .Name = "piyo",
        '            .DataTables = New List(Of DataTableModel) From {
        '                New DataTableModel With {.Name = "piyohoge"},
        '                New DataTableModel With {.Name = "piyofuga"},
        '                New DataTableModel With {.Name = "piyopiyo"}
        '            }
        '        }
        '    }
        '}
        'Me.Server = test
    End Sub

    Sub Test1(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        MsgBox("hoge")
    End Sub
    Sub Test2(ByVal sender As Object, ByVal e As NotifyCollectionChangedEventArgs)
        MsgBox("hoge")
    End Sub

    Sub New()
        Dim pm As New ProjectModel
        Dim proxy As ProjectCheckProxy
        Dim proxy2 As ModelLoadProxy(Of ServerModel)

        proxy = AddressOf pm.FileCheck
        proxy2 = AddressOf pm.ModelLoad(Of ServerModel)

        'Select Case pm.ProjectCheck(proxy, _DBEXPLORERVIEW_FILE)
        '    Case 0
        '        Me.Server = proxy2(_DBEXPLORERVIEW_FILE)
        '        If Me.Server IsNot Nothing Then

        '        Else
        '            'Call Me._SetDefault()
        '            Call Me._SetTest()
        '        End If
        '    Case Else
        '        'Call Me._SetDefault()
        '        Call Me._SetTest()
        'End Select
    End Sub
End Class
