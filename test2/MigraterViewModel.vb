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

    Sub Test()
        MsgBox("hoge")
    End Sub

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
                .FieldValue = "fuga",
                .Conditions = New ObservableCollection(Of ConditionModel) From {
                    New ConditionModel With {
                        .FieldName = "hogehoge",
                        .FieldValue = "fugafuga"
                    }
                }
            }
        }
        '------------------------------------------------------------------'

        ' コマンドフラグの有効／無効化
    End Sub
End Class
