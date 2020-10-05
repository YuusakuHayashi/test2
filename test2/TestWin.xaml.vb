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
        '   - AppDirectoryModel ... アプリケーションの構成を定義
        '   - ProjectInfoModel  ... プロジェクトの構成を定義
        '

        ' Of T型は何でも良い
        Dim ml As New ModelLoader(Of Nullable)

        ' AppDirectoryModelはここでロード
        Dim adm As AppDirectoryModel

        ' ProjectInfoModel, Model, ViewModelは各プログラムでロード（更新）
        Dim pim As New ProjectInfoModel
        Dim m As New Model
        Dim vm As New ViewModel

        adm = ml.ModelLoad(Of AppDirectoryModel)(AppDirectoryModel.ModelFileName)
        If adm Is Nothing Then
            adm = New AppDirectoryModel
        End If
        adm.ModelSave(AppDirectoryModel.ModelFileName, adm)

        'Dim udvm As New UserDirectoryViewModel(m, vm, adm, pim)
        Dim udvm As New UserDirectoryViewModel
        Call udvm.MyInitializing(m, vm, adm, pim)

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
