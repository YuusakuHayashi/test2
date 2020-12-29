Imports System.Windows.Forms

Public Module RpaModule
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

    Public Function LoadCurrentRpa(ByRef ini As Object) As Object
        If Not ini.AutoLoad Then
            Return Nothing
        End If
        If ini.CurrentSolution Is Nothing Then
            Return Nothing
        End If

        Dim rpa = RpaObject(ini.CurrentSolution.Architecture)
        rpa = rpa.Load(ini.CurrentSolution.JsonFileName)
        If rpa Is Nothing Then
            Console.WriteLine($"ファイル '{ini.CurrentSolution.JsonFileName}' のロードに失敗しました")
            Console.WriteLine()
            Return Nothing
        Else
            Console.WriteLine($"ファイル '{ini.CurrentSolution.JsonFileName}' をロードしました")
            Console.WriteLine()
            Return rpa
        End If
    End Function

    'Public Function ShowRpaIndicator(ByRef rpa As Object) As String
    '    If rpa IsNot Nothing Then
    '        Console.Write($"{rpa.SolutionName}\{rpa.ProjectAlias}>")
    '    Else
    '        Console.Write("NoRpa>")
    '    End If
    '    Return Console.ReadLine()
    'End Function

    'Public Function ShowRpaIndicator(ByRef rpa As Object, ByVal mode As String) As String
    '    If String.IsNullOrEmpty(mode) Then
    '        Return ShowRpaIndicator(rpa)
    '    Else
    '        Console.Write($"{mode}>")
    '        Return Console.ReadLine()
    '    End If
    'End Function

    Public Function SetDirectoryFromDialog(ByRef trn As RpaTransaction, ByRef rpa As Object, ByRef ini As RpaInitializer, ByVal propname As String) As String
        Dim [dir] As String = vbNullString
        Dim yorn As String = vbNullString
        Dim yorn2 As String = vbNullString
        Dim fbd As FolderBrowserDialog
        Do
            yorn = vbNullString
            Console.WriteLine($"'{propname}' の設定を行いますか (y/n)")
            yorn = trn.ShowRpaIndicator(rpa)
            Console.WriteLine(vbNullString)
        Loop Until yorn = "y" Or yorn = "n"
        If yorn = "y" Then
            Do
                yorn2 = vbNullString
                fbd = New FolderBrowserDialog With {
                    .Description = $"Select MyDirectory",
                    .RootFolder = Environment.SpecialFolder.Desktop,
                    .SelectedPath = Environment.SpecialFolder.Desktop,
                    .ShowNewFolderButton = True
                }
                If fbd.ShowDialog() = DialogResult.OK Then
                    Console.WriteLine($"よろしいですか？ '{fbd.SelectedPath}' (y/n)")
                    yorn2 = trn.ShowRpaIndicator(rpa)
                    Console.WriteLine(vbNullString)
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
            Console.WriteLine(vbNullString)
        End If
        Return [dir]
    End Function
End Module
