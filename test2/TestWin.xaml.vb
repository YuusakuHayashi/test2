Option Explicit On

Imports System.ComponentModel

Delegate Sub ViewControllerProxy(ByVal vm As Object)

Public Class TestWin
    Private _mnavi As NavigationService
    Private _enavi As NavigationService

    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' 実装メモ
        '   - ProjectModel2
        '     |
        '     +-- BaseViewModel2 ... Inortify 実装
        '         |
        '         +-- Each ViewModel
        '
        '   - AppDirectoryModel
        '     |
        '     +-- Project
        '         |
        '         +-- BaseViewModel2
        '             |
        '             +-- Each ViewModel
        '
        '   - BaseModel
        '         + ＩＮｏｒｔｉｆｙ実装
        '
        '   - BaseViewModel
        '         + ＩＮｏｒｔｉｆｙ実装
        '
        '   - Ｍｏｄｅｌ、ＶｉｅｗＭｏｄｅｌの実装は親クラスから継承すること。
        '   - コマンド（ＩＣＯＭＭＡＮＤ）のプロパティは
        '     ＶｉｅｗＭｏｄｅｌを実装するクラスのみ実装すること。
        '     つまり、Ｍｏｄｅｌにはコマンドを実装しないこと
        '     また、ＶｉｅｗＭｏｄｅｌはコマンドを実装するため、

        Dim m As New Model2
        Dim vm As New ViewModel
        Dim pim As New ProjectInfoModel

        Dim udvm As New UserDirectoryViewModel(m, vm, pim)

        Me.MainFlame.DataContext = vm


        '' 可変ファイル名のロード
        'Dim fmm As New FileManagerModel
        'fmm = fmm.ModelLoad(fmm.FileManagerJson)
        'If fmm Is Nothing Then
        '    fmm = New FileManagerModel
        'End If
        'fmm.MemberCheck()
        'fmm.ModelSave(fmm.FileManagerJson, fmm)


        '' モデルのロード
        'Dim m As New Model
        '' fmm.ModelLoadと違い、こちらはオーバロードしている
        'm = m.ModelLoad(fmm.CurrentModelJson)
        'If m Is Nothing Then
        '    m = New Model
        'End If
        'm.MemberCheck()
        ''m.ModelSave(m.CurrentModelJson, m)



        'Dim vm As New ViewModel

        'Me.DataContext = vm

        '' 呼び出し順で優先度を変える。上に行くほど優先
        'Dim ivm As New InitViewModel(m, vm)
        'Dim mvm As New MigraterViewModel(m, vm)
        'Me.MainFlame.DataContext = vm

        'Dim dbevm As New DBExplorerViewModel(m, vm)
        'Me.ExplorerFlame.DataContext = vm

        'Dim hvm As New HistoryViewModel(m, vm)
        'Me.HistoryFlame.DataContext = vm

        'm.MenuFolder.MemberCheck(vm)
    End Sub
End Class
