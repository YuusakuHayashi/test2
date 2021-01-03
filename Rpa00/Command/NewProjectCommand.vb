Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms

Public Class NewProjectCommand : Inherits RpaCommandBase
    Public Overrides Function Execute(ByRef dat As RpaDataWrapper) As Integer
        Dim rpa As Object

        ' Solutionオブジェクト
        rpa = _SelectProjectArchitecture(dat)
        If rpa Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine()
            dat.Transaction.ExitFlag = True
            Return 1000
        End If

        ' ソリューション名の更新
        rpa = _GetSolution(dat, rpa)
        If rpa Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine()
            dat.Transaction.ExitFlag = True
            Return 1000
        End If

        rpa.MyDirectory = RpaModule.SetDirectoryFromDialog(dat, "MyDirectory")

        ' ソリューションをイニシャライザーへ登録
        Dim ini As RpaInitializer = _RegisterSolution(dat, rpa)
        If ini Is Nothing Then
            Console.WriteLine($"プロジェクトの新規作成を中止しました")
            Console.WriteLine()
            dat.Transaction.ExitFlag = True
            Return 1000
        End If

        Dim yorn As String = vbNullString
        If dat.Project IsNot Nothing Then
            Console.WriteLine($"現在起動しているソリューション '{dat.Project.ProjectName}' から切り替えますか？ (y/n)")
            Do
                yorn = vbNullString
                yorn = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
            Loop Until yorn = "y" Or yorn = "n"
        Else
            yorn = "y"
        End If
        If yorn = "y" Then
            dat.Project = rpa
        End If

        Directory.CreateDirectory(rpa.SystemProjectDirectory)
        Console.WriteLine($"ディレクトリ '{rpa.SystemProjectDirectory}' を新規作成しました")
        Call rpa.Save(rpa.SystemJsonFileName, rpa)
        Console.WriteLine($"ファイル     '{rpa.SystemJsonFileName}' を新規作成しました")
        Call ini.Save(RpaInitializer.SystemIniFileName, ini)
        Console.WriteLine($"ファイル     '{RpaInitializer.SystemIniFileName}' を更新しました")
        Console.WriteLine()
        Return 0
    End Function

    Private Function _SelectProjectArchitecture(ByRef dat As RpaDataWrapper) As Object
        Dim rpa As Object
        Dim yorn = vbNullString
        Do
            Dim eidx As Integer
            Dim eflg As Boolean
            Dim iidx As Integer
            Dim inp As String
            Do
                iidx = -1 : eidx = -1
                Console.WriteLine($"プロジェクト構成のインデックスを指定")
                Dim idx As Integer = 1
                eflg = True
                Do
                    Dim rpa2 As Object = RpaModule.RpaObject(idx)
                    If rpa2 IsNot Nothing Then
                        Console.WriteLine($"{idx} ... {rpa2.SystemArchTypeName}")
                    Else
                        Console.WriteLine($"{idx} ... やっぱりやめる")
                        eidx = idx
                        eflg = False
                    End If
                    idx += 1
                Loop Until (Not eflg)

                inp = dat.Transaction.ShowRpaIndicator(Nothing)
                If IsNumeric(inp) Then
                    iidx = Integer.Parse(inp)
                Else
                    iidx = eidx + 1
                End If
                Console.WriteLine()
            Loop Until iidx <= eidx

            rpa = RpaModule.RpaObject(iidx)

            If iidx = eidx Then
                rpa = Nothing
                Exit Do
            End If
            If rpa Is Nothing Then
                Console.WriteLine($"指定のインデックス '{inp}' からはオブジェクトを生成できませんでした")
                rpa = Nothing
                Exit Do
            End If

            If Not Directory.Exists(rpa.SystemArchDirectory) Then
                Directory.CreateDirectory(rpa.SystemArchDirectory)
            End If
            Do
                Console.WriteLine($"よろしいですか？ '{inp} ... {rpa.SystemArchTypeName}' (y/n)")
                yorn = dat.Transaction.ShowRpaIndicator(Nothing)
                Console.WriteLine()
            Loop Until yorn = "y" Or yorn = "n"
        Loop Until yorn = "y"

        Return rpa
    End Function

    ' rpaプロジェクトオブジェクトの[ProjectName]を更新する
    Private Function _GetSolution(ByRef dat As RpaDataWrapper, ByRef rpa As Object) As Object
        Dim [sln] As String = vbNullString
        Dim flag As Boolean = False
        Dim yorn As String = vbNullString

        Do
            flag = False
            [sln] = vbNullString

            Console.WriteLine($"プロジェクト名を入力してください ...")
            [sln] = dat.Transaction.ShowRpaIndicator(dat)
            rpa.ProjectName = [sln]
            Console.WriteLine()

            If Directory.Exists(rpa.SystemProjectDirectory) Then
                Console.WriteLine($"プロジェクト '{rpa.ProjectName}' は存在します")
                Console.WriteLine($"他のプロジェクト名を入力してください")
                Console.WriteLine()
                Continue Do
            End If

            Do
                yorn = vbNullString
                Console.WriteLine($"'{rpa.ProjectName}' よろしいですか？ (y/n)")
                yorn = dat.Transaction.ShowRpaIndicator(dat)
                Console.WriteLine()
            Loop Until yorn = "y" Or yorn = "n"

            If yorn = "y" Then
                flag = True
            Else
                rpa = Nothing
                flag = True
            End If
        Loop Until flag

        Return rpa
    End Function

    Private Function _RegisterSolution(ByRef dat As RpaDataWrapper, ByRef rpa As Object) As RpaInitializer
        Dim sl As RpaInitializer.RpaProject       ' 既にあるソリューション
        Dim [new] As RpaInitializer.RpaProject    ' 新規作成するソリューション
        Dim idx As Integer = -1
        Dim [sln] As String = rpa.ProjectName

        sl = dat.Initializer.Projects.Find(
            Function(s)
                Return (s.Name = [sln])
            End Function
        )
        [new] = New RpaInitializer.RpaProject With {
            .Name = rpa.ProjectName,
            .Architecture = rpa.SystemArchType,
            .ProjectDirectory = rpa.SystemProjectDirectory,
            .JsonFileName = rpa.SystemJsonFileName
        }
        If sl Is Nothing Then
            dat.Initializer.Projects.Add([new])
        Else
            idx = dat.Initializer.Projects.IndexOf(sl)
            dat.Initializer.Projects(idx) = [new]
        End If
        dat.Initializer.CurrentProject = [new]

        Return dat.Initializer
    End Function
End Class
