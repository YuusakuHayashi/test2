Public Class TreeViewModel : Inherits BaseViewModel

    Private _Model As TreeModel
    Public Property Model As TreeModel
        Get
            Return Me._Model
        End Get
        Set(value As TreeModel)
            Me._Model = value
        End Set
    End Property

    'Public Property Model As TreeModel


    Sub New()
        'Dim tm As TreeModel
        'tm = New TreeModel("Root") From {
        '    New TreeModel("Node1") From {
        '        New TreeModel("Node1-1"),
        '        New TreeModel("Node1-2") From {
        '            New TreeModel("Node1-2-1"),
        '            New TreeModel("Node1-2-2")
        '        }
        '    },
        '    New TreeModel("Node2") From {
        '        New TreeModel("Node2-1") From {
        '            New TreeModel("Node2-1-1"),
        '            New TreeModel("Node2-1-2")
        '        }
        '    }
        '}
        'Dim pm As New ProjectModel
        'Dim proxy As ModelSaveProxy(Of List(Of SaveTreeModel))
        'Dim fn As String : fn = Environment.GetEnvironmentVariable("USERPROFILE") & "\test\ConfigFile.json"
        'proxy = AddressOf pm.ModelSave(Of List(Of SaveTreeModel))
        'Call proxy(fn, ConvertSaveModel(tm))

        'Dim pm As New ProjectModel
        'Dim proxy As ModelLoadProxy(Of List(Of SaveTreeModel))
        'Dim fn As String : fn = Environment.GetEnvironmentVariable("USERPROFILE") & "\test\ConfigFile.json"
        'proxy = AddressOf pm.ModelLoad(Of List(Of SaveTreeModel))
        'Dim lstm As List(Of SaveTreeModel)
        'lstm = proxy(fn)
        'Me.Model = ConvertTreeModel(lstm)
    End Sub


    Function ConvertTreeModel(lstm As List(Of SaveTreeModel),
                              Optional tm As TreeModel = Nothing) As TreeModel

        Dim youngest As TreeModel

        If tm Is Nothing Then
            tm = New TreeModel("Root")
        End If
        If lstm IsNot Nothing Then
            For Each c In lstm
                tm.Children.Add(New TreeModel(c.RealName))
                If c.Children IsNot Nothing Then
                    youngest = tm.Children.Last
                    ConvertTreeModel(c.Children, youngest)
                End If
            Next
        End If
        ConvertTreeModel = tm
    End Function



    Function ConvertSaveModel(tm As TreeModel) As List(Of SaveTreeModel)
        Dim lstm As New List(Of SaveTreeModel)
        Dim stm As SaveTreeModel

        If tm IsNot Nothing Then
            For Each c In tm
                stm = New SaveTreeModel
                stm.RealName = c.RealName
                If c.Children.Count > 0 Then
                    stm.Children = New List(Of SaveTreeModel)
                    For Each c2 In ConvertSaveModel(c)
                        stm.Children.Add(c2)
                    Next
                End If
                lstm.Add(stm)
            Next
        End If
        ConvertSaveModel = lstm
    End Function


    '子孫を全列挙
    Private Function AllNodes(node As TreeModel) As IEnumerable(Of TreeModel)
        If node Is Nothing Then
            Return Enumerable.Empty(Of TreeModel)
        End If
        Return node.Concat(node.SelectMany(Function(x) AllNodes(x)))
    End Function

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
