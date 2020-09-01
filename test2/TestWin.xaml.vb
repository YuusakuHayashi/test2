﻿Imports System.ComponentModel
Public Class TestWin
    Private _mnavi As NavigationService
    Private _enavi As NavigationService

    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。


        ' Delegater
        Dim loader As ModelLoadProxy1(Of Model)
        Dim loader2 As ModelLoadProxy1(Of FileManagerModel)
        Dim saver As ModelSaveProxy1(Of Model)
        Dim saver2 As ModelSaveProxy1(Of FileManagerModel)

        ' Models
        Dim m As Model
        Dim pm As New ProjectModel
        Dim fmm As FileManagerModel

        ' ViewModels
        Dim ivm As InitViewModel
        Dim dbevm As DBExplorerViewModel
        Dim mvm As MigraterViewModel
        'Dim mvm As MenuViewModel

        ' Views
        Dim iv As InitView
        Dim dbev As DBExplorerView
        Dim mv As MigraterView
        'Dim mv As MenuView


        Me._mnavi = Me.MainFlame.NavigationService
        Me._enavi = Me.ExplorerFlame.NavigationService

        loader = AddressOf pm.ModelLoad(Of Model)
        loader2 = AddressOf pm.ModelLoad(Of FileManagerModel)
        saver = AddressOf pm.ModelSave(Of Model)
        saver2 = AddressOf pm.ModelSave(Of FileManagerModel)


        ' ファイル管理モデルのロード
        fmm = New FileManagerModel
        fmm = loader2(pm.FileManagerFileName)

        ' なければＪＳＯＮをとりあえず作っておく
        If fmm Is Nothing Then
            fmm = New FileManagerModel
            fmm.LatestSettingFileName = pm.DefaultSettingFileName
            Call saver2(pm.FileManagerFileName, fmm)
        Else
            If String.IsNullOrEmpty(fmm.LatestSettingFileName) Then
                fmm.LatestSettingFileName = pm.DefaultSettingFileName
                Call saver2(pm.FileManagerFileName, fmm)
            End If
        End If


        ' モデルのロード
        m = loader(fmm.LatestSettingFileName)

        ' なければＪＳＯＮをとりあえず作っておく
        If m Is Nothing Then
            m = New Model
            Call saver(fmm.LatestSettingFileName, m)
        End If


        m.SourceFile = fmm.LatestSettingFileName

        ' メイン画面ビューモデル
        ivm = New InitViewModel(m)
        mvm = New MigraterViewModel(m)

        ' 初期メイン画面
        If ivm.NextFlag Then
            mv = New MigraterView(mvm)
            Me._mnavi.Navigate(mv)
        Else
            iv = New InitView(ivm)
            Me._mnavi.Navigate(iv)
        End If


        ' エクスプローラー画面ビューモデル
        dbevm = New DBExplorerViewModel(m)

        ' 初期エクスプローラー画面
        dbev = New DBExplorerView(dbevm)
        Me._enavi.Navigate(dbev)



        AddHandler m.PropertyChanged, AddressOf Me._PageNavigation
        'AddHandler mvm.PropertyChanged, AddressOf Me._PageNavigation
    End Sub

    Private Sub _PageNavigation(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        Dim m As Model
        Dim pm As New ProjectModel
        Dim saver As ModelSaveProxy1(Of Model)

        saver = AddressOf pm.ModelSave(Of Model)

        If e.PropertyName = "ChangePageStrings" Then
            m = CType(sender, Model)
            saver(m.SourceFile, m)
        End If

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
