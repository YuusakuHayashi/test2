Imports System.IO
Imports System.Windows.Forms

Public Module RpaModule
    ' ファイルの存在確認応答をループして行う
    Public Function FileCheckLoop(ByVal [file] As String, ByRef dat As RpaDataWrapper) As Boolean
        Do
            Console.WriteLine($"以下のファイルを用意したら、キーを押してください・・・")
            Console.WriteLine($"ファイル名 : '{[file]}'")
            dat.Transaction.ShowRpaIndicator(dat)
            Console.WriteLine()
            If IO.File.Exists([file]) Then
                Return True
            Else
                Dim yorn As String = vbNullString
                Do
                    yorn = vbNullString
                    Console.WriteLine($"ファイル '{[file]}' が存在しません。")
                    Console.WriteLine($"'y'...再度確認  'n'...確認しないでチェックを終了")
                    yorn = dat.Transaction.ShowRpaIndicator(dat)
                    Console.WriteLine()
                Loop Until yorn = "y" Or yorn = "n"
                If yorn = "n" Then
                    Return False
                End If
            End If
        Loop Until True
        Return False
    End Function

    Public Function Pop(Of T As {New, IList})(ByVal [old] As T) As T
        Dim [new] As New T
        Dim i As Integer = 0
        For Each elm In old
            If i <> 0 Then
                [new].Add(elm)
            End If
            i += 1
        Next
        Return [new]
    End Function

    Public Const DEFUALTENCODING As String = "Shift-JIS"

    Public Sub Save(ByVal savefile As String, ByRef obj As Object, ByVal chgfile As String)
        obj.Save(savefile, obj)
        File.Delete(chgfile)
        Console.WriteLine($"ファイル '{savefile} をセーブしました")
        Console.WriteLine()
    End Sub

    Public Sub CreateChangedFile(ByVal f As String)
        Dim sw As System.IO.StreamWriter
        sw = New System.IO.StreamWriter(
            f, False, System.Text.Encoding.GetEncoding(RpaModule.DEFUALTENCODING)
        )
        sw.Close()
        sw.Dispose()
    End Sub


    Public ReadOnly Property RpaObject(ByVal index As Integer) As Object
        Get
            Dim ics As New IntranetClientServerProject
            Dim sap As New StandAloneProject
            Dim csp As New ClientServerProject
            Select Case index
                Case ics.SystemArchType : Return ics
                Case sap.SystemArchType : Return sap
                Case csp.SystemArchType : Return csp
                Case Else : Return Nothing
            End Select
        End Get
    End Property

    Public Function LoadCurrentRpa(ByRef dat As Object) As Object
        If Not dat.Initializer.AutoLoad Then
            Return Nothing
        End If
        If dat.Initializer.CurrentProject Is Nothing Then
            Console.WriteLine($"CurrentProject がないため、ロードしませんでした")
            Console.WriteLine()
            Return Nothing
        End If

        Dim rpa = RpaObject(dat.Initializer.CurrentProject.Architecture)
        rpa = rpa.Load(dat.Initializer.CurrentProject.JsonFileName)
        If rpa Is Nothing Then
            Console.WriteLine($"ファイル '{dat.Initializer.CurrentProject.JsonFileName}' のロードに失敗しました")
            Console.WriteLine()
            Return Nothing
        Else
            Console.WriteLine($"ファイル '{dat.Initializer.CurrentProject.JsonFileName}' をロードしました")
            Console.WriteLine()
            Return rpa
        End If
    End Function

    Public Function SetDirectoryFromDialog(ByRef dat As RpaDataWrapper, ByVal propname As String) As String
        Dim [dir] As String = vbNullString
        Dim yorn As String = vbNullString
        Dim yorn2 As String = vbNullString
        Dim fbd As FolderBrowserDialog
        Do
            yorn = vbNullString
            Console.WriteLine()
            Console.WriteLine($"'{propname}' を指定しますか？ (y/n)")
            yorn = dat.Transaction.ShowRpaIndicator(dat)
        Loop Until yorn = "y" Or yorn = "n"

        If yorn = "y" Then
            Do
                yorn2 = vbNullString
                fbd = New FolderBrowserDialog With {
                    .Description = $"Select {propname}",
                    .RootFolder = Environment.SpecialFolder.Desktop,
                    .SelectedPath = Environment.SpecialFolder.Desktop,
                    .ShowNewFolderButton = True
                }
                If fbd.ShowDialog() = DialogResult.OK Then
                    Console.WriteLine()
                    Console.WriteLine($"よろしいですか？ '{fbd.SelectedPath}' (y/n)")
                    yorn2 = dat.Transaction.ShowRpaIndicator(dat)
                Else
                    yorn2 = vbNullString
                End If
            Loop Until yorn2 = "y" Or yorn2 = vbNullString

            If yorn2 = "y" Then
                [dir] = fbd.SelectedPath
                Console.WriteLine()
                Console.WriteLine($"'{propname}' が設定されました")
            Else
                Console.WriteLine($"'{propname}' の設定は行いませんでした")
            End If
            Console.WriteLine()
        End If

        Return [dir]
    End Function

    Public Function SetFileFromDialog(ByRef dat As RpaDataWrapper, ByVal propname As String) As String
        Dim [file] As String = vbNullString
        Dim yorn As String = vbNullString
        Dim yorn2 As String = vbNullString
        Dim ofd As OpenFileDialog
        Do
            yorn = vbNullString
            Console.WriteLine()
            Console.WriteLine($"'{propname}' を指定しますか？ (y/n)")
            yorn = dat.Transaction.ShowRpaIndicator(dat)
        Loop Until yorn = "y" Or yorn = "n"

        If yorn = "y" Then
            Do
                yorn2 = vbNullString
                ofd = New OpenFileDialog With {
                    .FileName = "hoge.json",
                    .InitialDirectory = $"{CommonProject.SystemDirectory}",
                    .Filter = "JSONファイル(*.json)|*.json|すべてのファイル(*.*)|*.*",
                    .FilterIndex = 1,
                    .Title = $"Select {propname}",
                    .RestoreDirectory = True
                }
                If ofd.ShowDialog() = DialogResult.OK Then
                    Console.WriteLine()
                    Console.WriteLine($"よろしいですか？ '{ofd.FileName}' (y/n)")
                    yorn2 = dat.Transaction.ShowRpaIndicator(dat)
                Else
                yorn2 = vbNullString
                End If
        Loop Until yorn2 = "y" Or yorn2 = vbNullString

        If yorn2 = "y" Then
            [file] = ofd.FileName
            Console.WriteLine()
            Console.WriteLine($"'{propname}' が指定されました")
        Else
            Console.WriteLine($"'{propname}' が指定されませんでした")
        End If
            Console.WriteLine()
        End If

        Return [file]
    End Function
End Module
