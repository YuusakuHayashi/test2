﻿Imports System.IO

Public Class ProjectInfoModel
    Inherits JsonHandler(Of ProjectInfoModel)

    Private Const SHIFT_JIS As String = "Shift-JIS"

    Private _Name As String
    Public Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            _Name = value
        End Set
    End Property

    Private _Kind As String
    Public Property Kind As String
        Get
            Return _Kind
        End Get
        Set(value As String)
            _Kind = value
        End Set
    End Property

    Private _DirectoryName As String
    Public Property DirectoryName As String
        Get
            Return _DirectoryName
        End Get
        Set(value As String)
            _DirectoryName = value
        End Set
    End Property

    ' プロジェクトのモデルファイルです
    Public ReadOnly Property ModelFileName As String
        Get
            Return DirectoryName & "\Model.json"
        End Get
    End Property

    ' プロジェクトがアプリケーションにより作成されたことを表すファイルです
    Public ReadOnly Property IniFileName As String
        Get
            Return DirectoryName & "\Project.ini"
        End Get
    End Property

    Public ReadOnly Property ProjectInfoFileName As String
        Get
            Return DirectoryName & "\ProjectInfo.json"
        End Get
    End Property

    Public Delegate Sub ProjectLaunchProxy()

    Private Sub CreateProjectDirectory()
        Try
            Directory.CreateDirectory(Me.DirectoryName)
        Catch ex As Exception
            Throw New Exception
        End Try
    End Sub

    Public Sub CreateFile(ByVal f As String)
        Dim txt As String : txt = vbNullString
        Dim sw As System.IO.StreamWriter

        Try
            sw = New System.IO.StreamWriter(
                f, False, System.Text.Encoding.GetEncoding(SHIFT_JIS)
            )
            sw.Write(txt)
        Catch ex As Exception
            '
        Finally
            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
            End If
        End Try
    End Sub

    Private Sub CreateProjectModelFile()
        Call CreateFile(Me.ModelFileName)
    End Sub

    Private Sub CreateProjectIniFile()
        Call CreateFile(Me.IniFileName)
    End Sub

    Private Sub CreateProjectInfoFile()
        Call CreateFile(Me.ProjectInfoFileName)
    End Sub

    Public Overloads Function LoadProject(ByVal p As ProjectInfoModel) As ProjectInfoModel
        Dim project As ProjectInfoModel
        project = p.ModelLoad(p.ProjectInfoFileName)
        LoadProject = project
    End Function

    Public Overloads Function LoadProject(ByVal f As String) As ProjectInfoModel
        Dim project As ProjectInfoModel
        Dim ml As New JsonHandler(Of Nullable)
        project = ml.ModelLoad(Of ProjectInfoModel)(f)
        LoadProject = project
    End Function

    ' この関数はプロジェクトディレクトリの存在チェックを行います
    Public Function CheckStructure() As Integer
        Dim code As Integer : code = -1
        If Directory.Exists(Me.DirectoryName) Then
            If File.Exists(Me.ProjectInfoFileName) Then
                If File.Exists(Me.IniFileName) Then
                    code = 0
                Else
                    code = 10
                End If
            Else
                code = 100
            End If
        Else
            code = 1000
        End If
        CheckStructure = code
    End Function

    'Public Function CheckProjectDirectory() As Integer
    '    Dim code As Integer : code = 99
    '    If Directory.Exists(Me.DirectoryName) Then
    '        code = 10
    '        If File.Exists(Me.ProjectInfoFileName) Then
    '            code = 20
    '            If File.Exists(Me.IniFileName) Then
    '                code = 0
    '            End If
    '        End If
    '    End If
    '    CheckProjectDirectory = code
    'End Function


    ' この関数はプロジェクトの作成を行い、その結果を返します
    Public Sub Launch()
        Call Me.CreateProjectDirectory()
        Call Me.CreateProjectIniFile()
        Call Me.CreateProjectInfoFile()
        Call Me.CreateProjectModelFile()
    End Sub

    'Public Function ProjectLaunch() As Integer
    '    Dim proxy(3) As ProjectLaunchProxy
    '    Dim proxy2 As ProjectLaunchProxy

    '    proxy(0) = AddressOf CreateProjectDirectory
    '    proxy(1) = AddressOf CreateProjectIniFile
    '    proxy(2) = AddressOf CreateProjectModelFile
    '    proxy(3) = AddressOf CreateProjectInfoFile

    '    Dim code As Integer : code = CheckProjectDirectory()
    '    Try
    '        Select Case code
    '            Case 0
    '            Case 10
    '            Case 99
    '                proxy2 = [Delegate].Combine(proxy)
    '            Case Else
    '        End Select
    '        If proxy2 IsNot Nothing Then
    '            If proxy2.GetInvocationList IsNot Nothing Then
    '                Call proxy2()
    '                code = 0
    '            End If
    '        End If
    '    Catch ex As Exception
    '    Finally
    '        ProjectLaunch = code
    '    End Try
    'End Function
End Class