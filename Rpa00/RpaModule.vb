Imports System.Windows.Forms

Public Module RpaModule
    Public Const DEFUALTENCODING As String = "Shift-JIS"

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
            Console.WriteLine($"'{propname}' の設定を行いますか (y/n)")
            Console.WriteLine()
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
                    Console.WriteLine($"よろしいですか？ '{fbd.SelectedPath}' (y/n)")
                    Console.WriteLine()
                    yorn2 = dat.Transaction.ShowRpaIndicator(dat)
                Else
                    yorn2 = vbNullString
                End If
            Loop Until yorn2 = "y" Or yorn2 = vbNullString

            If yorn2 = "y" Then
                [dir] = fbd.SelectedPath
                Console.WriteLine($"'{propname}' が設定されました")
            Else
                Console.WriteLine($"'{propname}' の設定は行いませんでした")
            End If
            Console.WriteLine()
        End If

        Return [dir]
    End Function
End Module
