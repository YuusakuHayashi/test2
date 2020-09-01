Imports System.Collections.ObjectModel

Public Class MigraterViewModel : Inherits ViewModel
    ' モデル
    Private __Model As Model
    Private Property _Model As Model
        Get
            Return Me.__Model
        End Get
        Set(value As Model)
            Me.__Model = value
            RaisePropertyChanged("Model")
        End Set
    End Property


    Private _Conditions As ObservableCollection(Of ConditionModel)
    Public Property Conditions As ObservableCollection(Of ConditionModel)
        Get
            Return Me._Conditions
        End Get
        Set(value As ObservableCollection(Of ConditionModel))
            Me._Conditions = value
        End Set
    End Property

    ' コマンドプロパティ(リロード)
    Private _OrCommand As ICommand
    Public ReadOnly Property OrCommand As ICommand
        Get
            If Me._OrCommand Is Nothing Then
                Me._OrCommand = New DelegateCommand With {
                    .ExecuteHandler = AddressOf _OrCommandExecute,
                    .CanExecuteHandler = AddressOf _OrCommandCanExecute
                }
                Return Me._OrCommand
            Else
                Return Me._OrCommand
            End If
        End Get
    End Property

    ' コマンド有効／無効判定(リロード)
    Private Sub _CheckOrEnableFlag()
        'Dim b As Boolean
        'If String.IsNullOrEmpty(Me._Model.ConnectionString) Then
        '    b = False
        'Else
        '    b = True
        'End If
        'Me._OrEnableFlag = b
    End Sub

    ' コマンド実行(リロード)
    Private Sub _OrCommandExecute(ByVal parameter As Object)
        MsgBox("Hello Test")
    End Sub

    ' コマンド有効／無効化(リロード)
    Private Function _OrCommandCanExecute(ByVal parameter As Object) As Boolean
        'Return Me._OrEnableFlag
        Return True
    End Function


    Sub New(ByRef m As Model)
        ' モデルの設定
        If m IsNot Nothing Then
        Else
            ' モデルなし
            m = New Model
        End If
        Me._Model = m


        ' ビューモデルの設定
        ' TEST ------------------------------------------------------------'
        Me.Conditions = New ObservableCollection(Of ConditionModel) From {
            New ConditionModel With {
                .FieldName = "hoge",
                .FieldValue = "fuga"
            }
        }
        '------------------------------------------------------------------'

        ' コマンドフラグの有効／無効化
    End Sub
End Class
