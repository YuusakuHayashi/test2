Imports System.ComponentModel
Imports System.Collections.ObjectModel


Public Class TreeViewVM
    Inherits ViewModel

    Private _Datas As List(Of TreeViewModel)
    Public Property Datas As List(Of TreeViewModel)
        Get
            Return _Datas
        End Get
        Set(value As List(Of TreeViewModel))
            _Datas = value
        End Set
    End Property


    Sub New(ByRef tvm As List(Of TreeViewModel))

        Me.Datas = tvm

        'Dim pm As New ProjectModel
        'Me.Datas = New List(Of TreeViewModel)
        'Me.Datas = pm.LoadTreeViewModel
        'Me.Datas = New List(Of TreeViewM) From {
        '    New TreeViewM With {
        '        .RealName = "AAA", .Child = New List(Of TreeViewM) From {
        '            New TreeViewM With {
        '                .RealName = "ABA"
        '            },
        '            New TreeViewM With {
        '                .RealName = "ACA"
        '            }
        '        }
        '    },
        '    New TreeViewM With {
        '        .RealName = "BAA", .Child = New List(Of TreeViewM) From {
        '            New TreeViewM With {
        '                .RealName = "BBA"
        '            },
        '            New TreeViewM With {
        '                .RealName = "BCA"
        '            }
        '        }
        '    }
        '}
        'm.TreeViews = Me.Datas
        'm.ModelSave()
        'Dim s As ServerModel
        'Me._VM = New Model
        's = Me._VM.ServerLoad
        'Me.ServerVM = New List(Of ServerModel)
        'Me.ServerVM.Add(s)
    End Sub
End Class
