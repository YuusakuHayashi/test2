Public Class TreeViewModel : Inherits ViewModel

    Private _Model As TreeModel
    Public Property Model As TreeModel
        Get
            Return Me._Model
        End Get
        Set(value As TreeModel)
            Me._Model = value
        End Set
    End Property


    Sub New()
        Me.Model = New TreeModel("Root") From {
            New TreeModel("Node1") From {
                New TreeModel("Node1-1"),
                New TreeModel("Node1-2") From {
                    New TreeModel("Node1-2-1"),
                    New TreeModel("Node1-2-2")
                }
            },
            New TreeModel("Node2") From {
                New TreeModel("Node2-1") From {
                    New TreeModel("Node2-1-1"),
                    New TreeModel("Node2-1-2")
                }
            }
        }
    End Sub

    'TreeView は考え方が複雑・・・
    '描画線用のデータ構造ではTreeViewのリスト型

    'Public ServerName As String
    'Public DataBaseName As String
    'Public DataTableName As String

    'Private Property _RealName As String
    'Public Property RealName As String
    '    Get
    '        Return _RealName
    '    End Get
    '    Set(value As String)
    '        _RealName = value
    '        RaisePropertyChanged("RealName")
    '    End Set
    'End Property

    'Private Property _Checked As Boolean
    'Public Property Checked As Boolean
    '    Get
    '        Return _Checked
    '    End Get
    '    Set(value As Boolean)
    '        _Checked = value
    '        RaisePropertyChanged("Checked")
    '    End Set
    'End Property

    'Private _Child As List(Of TreeViewModel)
    'Public Property Child As List(Of TreeViewModel)
    '    Get
    '        Return _Child
    '    End Get
    '    Set(value As List(Of TreeViewModel))
    '        _Child = value
    '    End Set
    'End Property
    '-----------------------------------------------------------------------------'

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

    ''Private _AccessEnableFlag As Boolean
    ''Public Property AccessEnableFlag As Boolean
    ''    Get
    ''        Return _AccessEnableFlag
    ''    End Get
    ''    Set(value As Boolean)
    ''        _AccessEnableFlag = value
    ''        'RaisePropertyChanged("AccessEnableFlag")
    ''        CType(ClickCommand, DelegateCommand).RaiseCanExecuteChanged()
    ''    End Set
    ''End Property

    'Private Sub _ClickCommandExecute(ByVal parameter As Object)
    '    MsgBox("hello")
    'End Sub

    'Private Function _ClickCommandCanExecute(ByVal parameter As Object) As Boolean
    '    Return True
    'End Function
End Class
