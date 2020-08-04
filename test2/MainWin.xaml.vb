﻿Public Class MainWin
    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Dim cfvm As New ConfigFileVM
        Me.ConfigFileValue.DataContext = cfvm


        Dim pm As New ProjectModel
        Me.TreeView.DataContext = New TreeViewVM(pm.LoadTreeViewModel)


        '-- Sql 関連 ---------------------------------------------'
        Dim sm As New SqlModel : sm = pm.LoadSqlModel
        Dim ssvm As New SqlStatusVM(sm)
        Me.ServerNameValue.DataContext = ssvm
        Me.DataBaseValue.DataContext = ssvm
        Me.DataTabelValue.DataContext = ssvm
        Me.FieldNameValue.DataContext = ssvm
        Me.SourceValue.DataContext = ssvm
        Me.DistinationValue.DataContext = ssvm
        '---------------------------------------------------------'

        Dim abvm As New AccessButtonVM(sm, ssvm)
        Me.AccessButton.DataContext = abvm


        Me.SaveButton.DataContext = New SaveButtonVM(sm, ssvm, abvm)
    End Sub
End Class