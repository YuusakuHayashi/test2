Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports Newtonsoft.Json

Public Class GreenCalenderViewModel : Inherits RpaViewModelBase(Of GreenCalenderViewModel)
    Public Class CellData : Inherits ViewModelBase
        Private _Day As Date
        Public Property Day As Date
            Get
                Return Me._Day
            End Get
            Set(value As Date)
                Me._Day = value
                RaisePropertyChanged("Day")
            End Set
        End Property

        Private _Density As Integer
        Public Property Density As Integer
            Get
                Return Me._Density
            End Get
            Set(value As Integer)
                Me._Density = value
                RaisePropertyChanged("Density")
            End Set
        End Property

        Private _Row As Integer
        <JsonIgnore>
        Public Property Row As Integer
            Get
                Return Me._Row
            End Get
            Set(value As Integer)
                Me._Row = value
                RaisePropertyChanged("Row")
            End Set
        End Property

        Private _Column As Integer
        <JsonIgnore>
        Public Property Column As Integer
            Get
                Return Me._Column
            End Get
            Set(value As Integer)
                Me._Column = value
                RaisePropertyChanged("Column")
            End Set
        End Property
    End Class

    Public Class ColumnHeaderData : Inherits ViewModelBase
        Private _Week As String
        Public Property Week As String
            Get
                Return Me._Week
            End Get
            Set(value As String)
                Me._Week = value
                RaisePropertyChanged("Week")
            End Set
        End Property

        Private _Row As Integer
        Public Property Row As Integer
            Get
                Return Me._Row
            End Get
            Set(value As Integer)
                Me._Row = value
                RaisePropertyChanged("Row")
            End Set
        End Property
    End Class

    Public Class RowHeaderData : Inherits ViewModelBase
        Private _Month As Double
        Public Property Month As Double
            Get
                Return Me._Month
            End Get
            Set(value As Double)
                Me._Month = value
                RaisePropertyChanged("Month")
            End Set
        End Property

        Private _Column As Integer
        Public Property Column As Integer
            Get
                Return Me._Column
            End Get
            Set(value As Integer)
                Me._Column = value
                RaisePropertyChanged("Column")
            End Set
        End Property
    End Class

    Private _YearString As String
    <JsonIgnore>
    Public Property YearString As String
        Get
            Return Me._YearString
        End Get
        Set(value As String)
            Me._YearString = value
            RaisePropertyChanged("YearString")
        End Set
    End Property

    Private _Range As ObservableCollection(Of CellData)
    Public Property Range As ObservableCollection(Of CellData)
        Get
            If Me._Range Is Nothing Then
                Me._Range = New ObservableCollection(Of CellData)
            End If
            Return Me._Range
        End Get
        Set(value As ObservableCollection(Of CellData))
            Me._Range = value
            RaisePropertyChanged("Range")
        End Set
    End Property

    Private _RowHeaders As ObservableCollection(Of RowHeaderData)
    <JsonIgnore>
    Public Property RowHeaders As ObservableCollection(Of RowHeaderData)
        Get
            If Me._RowHeaders Is Nothing Then
                Me._RowHeaders = New ObservableCollection(Of RowHeaderData)
            End If
            Return Me._RowHeaders
        End Get
        Set(value As ObservableCollection(Of RowHeaderData))
            Me._RowHeaders = value
            RaisePropertyChanged("RowHeaders")
        End Set
    End Property

    Private _ColumnHeaders As ObservableCollection(Of ColumnHeaderData)
    <JsonIgnore>
    Public Property ColumnHeaders As ObservableCollection(Of ColumnHeaderData)
        Get
            If Me._ColumnHeaders Is Nothing Then
                Me._ColumnHeaders = New ObservableCollection(Of ColumnHeaderData)
            End If
            Return Me._ColumnHeaders
        End Get
        Set(value As ObservableCollection(Of ColumnHeaderData))
            Me._ColumnHeaders = value
            RaisePropertyChanged("ColumnHeaders")
        End Set
    End Property

    Public Sub SetupGreenCalender()
        Dim act1 As Action = AddressOf InitializeCalender
        Dim act2 As Action = AddressOf GoCalenderFoward
        Dim act3 As Action = AddressOf SetCellIndex
        Dim act4 As Action = AddressOf SetMonthRowHeader
        Dim act5 As Action = AddressOf SetWeekColumnHeader
        Dim act6 As Action = AddressOf SetYear
        Dim all As Action = [Delegate].Combine(act1, act2, act3, act4, act5, act6)
        Call all()
    End Sub

    Private Sub InitializeCalender()
        If Me.Range.Count > 0 Then
            Dim latest As Date = Me.Range.Last.Day
            Dim term As Integer = (DateTime.Today - latest).TotalDays

            If term < 366 Then
                Exit Sub
            End If
        End If

        Me.Range = New ObservableCollection(Of CellData)
        Dim today As Date = DateTime.Today
        For i As Integer = 366 To 0 Step -1
            Me.Range.Add(New CellData With {.Day = today.AddDays(-i)})
        Next
    End Sub

    Private Sub GoCalenderFoward()
        Dim latest As Date = Me.Range.Last.Day
        Dim term As Integer = (DateTime.Today - latest).TotalDays

        Dim i As Integer = term
        Do
            If i <= 0 Then
                Exit Do
            End If
            Me.Range.RemoveAt(0)
            Me.Range.Add(New CellData With {.Day = DateTime.Today.AddDays(-i + 1)})
            i -= 1
        Loop Until False
    End Sub

    Private Sub SetCellIndex()
        Dim r As Integer = 0
        Dim c As Integer = 0
        For Each cell In Me.Range
            cell.Row = r
            cell.Column = c

            r += 1
            If r > 6 Then
                r = 0
                c += 1
            End If
        Next
    End Sub

    Private Sub SetMonthRowHeader()
        Dim term As Integer = -1
        Dim premon As Integer = 0
        Dim col As Integer = 0

        Do
            Dim rhd As New RowHeaderData
            rhd.Column = col

            Dim mon As Double = 0.0
            Dim topmon As Integer = Me.Range(term + 1).Day.Month
            mon = topmon

            Dim nowmon As Integer = 0
            For j As Integer = 2 To 7
                Dim idx As Integer = term + j
                If idx <= (Me.Range.Count - 1) Then
                    nowmon = Me.Range(idx).Day.Month
                    If mon <> nowmon Then
                        mon += 0.5
                        Exit For
                    End If
                End If
            Next

            If nowmon <> topmon Then
                premon = nowmon
            Else
                If mon = premon Then
                    mon = 0.0
                Else
                    premon = nowmon
                End If
            End If

            rhd.Month = mon
            Me.RowHeaders.Add(rhd)

            col += 1
            term += 7
            If term >= 366 Then
                Exit Do
            End If
        Loop Until False
    End Sub

    Private Sub SetWeekColumnHeader()
        Me.ColumnHeaders.Add(New ColumnHeaderData With {
            .Row = 0,
            .Week = Strings.Left(Me.Range(0).Day.DayOfWeek.ToString, 3)
        })
        Me.ColumnHeaders.Add(New ColumnHeaderData With {
            .Row = 3,
            .Week = Strings.Left(Me.Range(3).Day.DayOfWeek.ToString, 3)
        })
        Me.ColumnHeaders.Add(New ColumnHeaderData With {
            .Row = 6,
            .Week = Strings.Left(Me.Range(6).Day.DayOfWeek.ToString, 3)
        })
    End Sub

    Private Sub SetYear()
        Dim tyear As Integer = Me.Range.First.Day.Year.ToString
        Dim nyear As Integer = Me.Range.Last.Day.Year.ToString
        Me.YearString = $"({tyear} ~ {nyear})"
    End Sub
End Class
