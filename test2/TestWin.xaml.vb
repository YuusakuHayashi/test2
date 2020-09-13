Option Explicit On

Imports System.ComponentModel

Delegate Sub ViewControllerProxy(ByVal vm As Object)

Public Class TestWin
    Private _mnavi As NavigationService
    Private _enavi As NavigationService

    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' Models
        Dim m As Model
        Dim fmm As New FileManagerModel
        fmm = fmm.ModelLoad(fmm.FileManagerJson)
        If fmm Is Nothing Then
            fmm = New FileManagerModel
        End If
        Call fmm.MemberCheck()
        Call fmm.ModelSave(fmm.FileManagerJson, fmm)
        ' 1. 初期状態の時
        '    各クラスのデフォルト設定ファイルパスをファイルマネージャに登録する
        '    そのファイルにNewしたオブジェクトをセーブする
        ' 2. 次回以降
        '    ファイルマネージャに登録されたファイルからオブジェクトをロードする


        ' モデルのロード
        m = New Model
        m.CurrentModelJson = fmm.CurrentModelJson
        m = m.ModelLoad(m.CurrentModelJson)
        If m Is Nothing Then
            m = New Model
        End If
        Call m.MemberCheck()
        Call m.ModelSave(fmm.CurrentModelJson, m)



        Dim vm1 As New ViewModel
        Dim vm2 As New ViewModel

        ' 呼び出し順で優先度を変える。下に行くほど優先
        Dim mvm As New MigraterViewModel(m, vm1)
        Dim ivm As New InitViewModel(m, vm1)

        Me.MainFlame.DataContext = vm1

        Dim dbevm As New DBExplorerViewModel(m, vm2)
        Me.ExplorerFlame.DataContext = vm2

        AddHandler vm1.PropertyChanged, AddressOf Me._PageShift
        AddHandler vm2.PropertyChanged, AddressOf Me._PageShift
    End Sub

    Private Sub _PageShift(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)

        Dim m As Model
        Dim vm As ViewModel
        Dim mvm As MigraterViewModel
        Dim ivm As InitViewModel

        Me.MainFlame.DataContext = Nothing

        If Not e.PropertyName = "ContextDistination" Then
            Exit Sub
        End If

        Select Case sender.GetType.Name
            Case "InitViewModel"
                ivm = CType(sender, InitViewModel)
                m = ivm.Model
                vm = ivm.ViewModel
            Case "MigraterViewModel"
                mvm = CType(sender, MigraterViewModel)
                m = mvm.Model
                vm = mvm.ViewModel
        End Select

        'Select Case sender.GetType.Name
        '    Case "InitViewModel"
        'End Select
        'Dim m As Model
        'Dim pm As New ProjectModel
        'Dim saver As ModelSaveProxy1(Of Model)

        'saver = AddressOf pm.ModelSave(Of Model)

        'If e.PropertyName = "ChangePageStrings" Then
        '    m = CType(sender, Model)
        '    saver(m.SourceFile, m)
        'End If

        'Select Case sender.GetType.Name
        '    Case "InitViewModel"
        '        If e.PropertyName = "InitFlag" Then
        '            ivm = CType(sender, InitViewModel)
        '            If ivm.InitFlag Then
        '                mvm = New MenuViewModel
        '                mv = New MenuView(mvm)
        '                Me._mnavi.Navigate(mv)
        '            End If
        '        End If
        '        'Case "MenuViewModel"
        '        '    If e.PropertyName = "MenuFlag" Then
        '        '        mvm = CType(sender, MenuViewModel)
        '        '        If mvm.MenuFlag Then
        '        '            bmv = New BatchMenuView
        '        '            Me._mnavi.Navigate(bmv)
        '        '        End If
        '        '    End If
        '    Case Else
        '        Exit Sub
        'End Select
    End Sub
End Class
