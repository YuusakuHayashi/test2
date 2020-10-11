Imports System.Collections.ObjectModel

Public Class MenuFolderModel
    Inherits BaseModel(Of MenuFolderModel)

    ' 親クラスからＶｉｅｗＭｏｄｅｌを継承すると、Ｓａｖｅが失敗するため

    Private __PrivateViewModel As ViewModel
    Private Property _PrivateViewModel As ViewModel
        Get
            Return __PrivateViewModel
        End Get
        Set(value As ViewModel)
            __PrivateViewModel = value
        End Set
    End Property


    Private _Menus As ObservableCollection(Of MenuModel)
    Public Property Menus As ObservableCollection(Of MenuModel)
        Get
            Return Me._Menus
        End Get
        Set(value As ObservableCollection(Of MenuModel))
            Me._Menus = value
            RaisePropertyChanged("Menus")
        End Set
    End Property


    Private _IsExpanded As Boolean
    Public Property IsExpanded As Boolean
        Get
            Return Me._IsExpanded
        End Get
        Set(value As Boolean)
            Me._IsExpanded = value
            RaisePropertyChanged("IsExpanded")
        End Set
    End Property


    ' コレクションの存在チェック
    Private Sub _MenusCheck()
        If Me.Menus Is Nothing Then
            Menus = New ObservableCollection(Of MenuModel)
        End If
    End Sub


    ' コレクションのセット
    Private Sub _MenusUpdate()
        Dim flg As Boolean
        For Each vdic In _PrivateViewModel.ContentDictionary
            For Each c In vdic.Value
                flg = True
                For Each m In Me.Menus
                    If m.Name = c.Key Then
                        flg = False
                    End If
                Next

                If flg Then
                    Me.Menus.Add(
                        New MenuModel With {
                            .ViewName = vdic.Key,
                            .Name = c.Key,
                            .DisplayName = c.Key
                        }
                    )
                End If
            Next
        Next
    End Sub


    ' 変更要求購読
    ' 購読をここで行うかは色々検討する点がある。
    Public Sub MenuChangeRequestedAddHandler()
        AddHandler _
            MenuChangedEventListener.Instance.ChangeRequested,
            AddressOf Me._MenuChangeRequestedReview
    End Sub


    ' 変更要求チェック
    Private Sub _MenuChangeRequestedReview(ByVal mm As MenuModel, ByVal e As System.EventArgs)
        Dim b As Boolean : b = True
        If b Then
            Me._MenuChangeRequestAccept(mm)
        End If
    End Sub


    ' 変更要求受理
    Private Sub _MenuChangeRequestAccept(ByVal mm As MenuModel)
        Me._PrivateViewModel.ChangeContent(mm.ViewName, mm.Name, mm)
    End Sub


    ' メンバーチェック
    Public Overloads Sub MemberCheck()
        Me._MenusCheck()
    End Sub


    ' メンバーチェック
    Public Overloads Sub MemberCheck(ByRef vm As ViewModel)
        Me._MenusCheck()
        Me._PrivateViewModel = vm
        Me._MenusUpdate()
    End Sub

    Sub New()
    End Sub
End Class
