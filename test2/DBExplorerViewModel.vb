Public Class DBExplorerViewModel
    Private Const _DBEXPLORERVIEW_FILE As String = "DBExplorerView.txt"

    Private _Server As ServerModel
    Public Property Server As ServerModel
        Get
            Return Me._Server
        End Get
        Set(value As ServerModel)
            Me._Server = value
        End Set
    End Property


    Sub _SetTest()
        '外部に記録
        Dim pm As New ProjectModel
        Dim proxy As ProjectCheckProxy
        proxy = AddressOf pm.FileCheck
        Select Case pm.ProjectCheck(proxy, _DBEXPLORERVIEW_FILE)
            Case 0
                pm.ModelSave(Of ServerModel)(_DBEXPLORERVIEW_FILE, Me.Server)
            Case 1
                pm.FileEstablish(_DBEXPLORERVIEW_FILE)
                pm.ModelSave(Of ServerModel)(_DBEXPLORERVIEW_FILE, Me.Server)
            Case 99
            Case 999
        End Select
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

    Sub New()
        Dim pm As New ProjectModel
        Dim proxy As ProjectCheckProxy
        Dim proxy2 As ModelLoadProxy(Of ServerModel)

        proxy = AddressOf pm.FileCheck
        proxy2 = AddressOf pm.ModelLoad(Of ServerModel)

        Select Case pm.ProjectCheck(proxy, _DBEXPLORERVIEW_FILE)
            Case 0
                Me.Server = proxy2(_DBEXPLORERVIEW_FILE)
                If Me.Server IsNot Nothing Then

                Else
                    'Call Me._SetDefault()
                    Call Me._SetTest()
                End If
            Case Else
                'Call Me._SetDefault()
                Call Me._SetTest()
        End Select
    End Sub
End Class
