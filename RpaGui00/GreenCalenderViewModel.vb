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

        Private _RowSpan As Integer
        <JsonIgnore>
        Public Property RowSpan As Integer
            Get
                Return Me._RowSpan
            End Get
            Set(value As Integer)
                Me._RowSpan = value
                RaisePropertyChanged("RowSpan")
            End Set
        End Property

        Private _ColumnSpan As Integer
        <JsonIgnore>
        Public Property ColumnSpan As Integer
            Get
                Return Me._ColumnSpan
            End Get
            Set(value As Integer)
                Me._ColumnSpan = value
                RaisePropertyChanged("ColumnSpan")
            End Set
        End Property
    End Class

    Public Class ColumnHeaderData : Inherits ViewModelBase
        Private _Header As String
        Public Property Header As String
            Get
                Return Me._Header
            End Get
            Set(value As String)
                Me._Header = value
                RaisePropertyChanged("Header")
            End Set
        End Property
        
        'Private _Week As String
        'Public Property Week As String
        '    Get
        '        Return Me._Week
        '    End Get
        '    Set(value As String)
        '        Me._Week = value
        '        RaisePropertyChanged("Week")
        '    End Set
        'End Property

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

        Private _RowSpan As Integer
        Public Property RowSpan As Integer
            Get
                Return Me._RowSpan
            End Get
            Set(value As Integer)
                Me._RowSpan = value
                RaisePropertyChanged("RowSpan")
            End Set
        End Property
    End Class

    Public Class RowHeaderData : Inherits ViewModelBase
        Private _Header As String
        Public Property Header As String
            Get
                Return Me._Header
            End Get
            Set(value As String)
                Me._Header = value
                RaisePropertyChanged("Header")
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

        Private _ColumnSpan As Integer
        Public Property ColumnSpan As Integer
            Get
                Return Me._ColumnSpan
            End Get
            Set(value As Integer)
                Me._ColumnSpan = value
                RaisePropertyChanged("ColumnSpan")
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

    Private _YearRange As ObservableCollection(Of CellData)
    Public Property YearRange As ObservableCollection(Of CellData)
        Get
            If Me._YearRange Is Nothing Then
                Me._YearRange = New ObservableCollection(Of CellData)
            End If
            Return Me._YearRange
        End Get
        Set(value As ObservableCollection(Of CellData))
            Me._YearRange = value
            RaisePropertyChanged("YearRange")
        End Set
    End Property

    Private _QuaterRange As ObservableCollection(Of CellData)
    <JsonIgnore>
    Public Property QuaterRange As ObservableCollection(Of CellData)
        Get
            If Me._QuaterRange Is Nothing Then
                Me._QuaterRange = New ObservableCollection(Of CellData)
            End If
            Return Me._QuaterRange
        End Get
        Set(value As ObservableCollection(Of CellData))
            Me._QuaterRange = value
            RaisePropertyChanged("QuaterRange")
        End Set
    End Property

    Private _MonthRange As ObservableCollection(Of CellData)
    <JsonIgnore>
    Public Property MonthRange As ObservableCollection(Of CellData)
        Get
            If Me._MonthRange Is Nothing Then
                Me._MonthRange = New ObservableCollection(Of CellData)
            End If
            Return Me._MonthRange
        End Get
        Set(value As ObservableCollection(Of CellData))
            Me._MonthRange = value
            RaisePropertyChanged("MonthRange")
        End Set
    End Property

    Private _WeekRange As ObservableCollection(Of CellData)
    <JsonIgnore>
    Public Property WeekRange As ObservableCollection(Of CellData)
        Get
            If Me._WeekRange Is Nothing Then
                Me._WeekRange = New ObservableCollection(Of CellData)
            End If
            Return Me._WeekRange
        End Get
        Set(value As ObservableCollection(Of CellData))
            Me._WeekRange = value
            RaisePropertyChanged("WeekRange")
        End Set
    End Property

    Private _Range As ObservableCollection(Of CellData)
    <JsonIgnore>
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

    Private _YearRowHeaders As ObservableCollection(Of RowHeaderData)
    Private Property YearRowHeaders As ObservableCollection(Of RowHeaderData)
        Get
            If Me._YearRowHeaders Is Nothing Then
                Me._YearRowHeaders = New ObservableCollection(Of RowHeaderData)
            End If
            Return Me._YearRowHeaders
        End Get
        Set(value As ObservableCollection(Of RowHeaderData))
            Me._YearRowHeaders = value
        End Set
    End Property

    Private _YearColumnHeaders As ObservableCollection(Of ColumnHeaderData)
    Private Property YearColumnHeaders As ObservableCollection(Of ColumnHeaderData)
        Get
            If Me._YearColumnHeaders Is Nothing Then
                Me._YearColumnHeaders = New ObservableCollection(Of ColumnHeaderData)
            End If
            Return Me._YearColumnHeaders
        End Get
        Set(value As ObservableCollection(Of ColumnHeaderData))
            Me._YearColumnHeaders = value
        End Set
    End Property

    Private _QuaterRowHeaders As ObservableCollection(Of RowHeaderData)
    Private Property QuaterRowHeaders As ObservableCollection(Of RowHeaderData)
        Get
            If Me._QuaterRowHeaders Is Nothing Then
                Me._QuaterRowHeaders = New ObservableCollection(Of RowHeaderData)
            End If
            Return Me._QuaterRowHeaders
        End Get
        Set(value As ObservableCollection(Of RowHeaderData))
            Me._QuaterRowHeaders = value
        End Set
    End Property

    Private _QuaterColumnHeaders As ObservableCollection(Of ColumnHeaderData)
    Private Property QuaterColumnHeaders As ObservableCollection(Of ColumnHeaderData)
        Get
            If Me._QuaterColumnHeaders Is Nothing Then
                Me._QuaterColumnHeaders = New ObservableCollection(Of ColumnHeaderData)
            End If
            Return Me._QuaterColumnHeaders
        End Get
        Set(value As ObservableCollection(Of ColumnHeaderData))
            Me._QuaterColumnHeaders = value
        End Set
    End Property

    Private _TermTypes As ObservableCollection(Of String)
    <JsonIgnore>
    Public ReadOnly Property TermTypes As ObservableCollection(Of String)
        Get
            If Me._TermTypes Is Nothing Then
                Me._TermTypes = New ObservableCollection(Of String)
                Me._TermTypes.Add("Year")
                Me._TermTypes.Add("Quater")
                Me._TermTypes.Add("Month")
                Me._TermTypes.Add("Week")
            End If
            Return Me._TermTypes
        End Get
    End Property

    Private _IsInitialized As Boolean
    Private Property IsInitialized As Boolean
        Get
            Return Me._IsInitialized
        End Get
        Set(value As Boolean)
            Me._IsInitialized = value
        End Set
    End Property

    Private _TermIndex As Integer
    <JsonIgnore>
    Public Property TermIndex As Integer
        Get
            Return Me._TermIndex
        End Get
        Set(value As Integer)
            Me._TermIndex = value
            RaisePropertyChanged("TermIndex")
            Call SetupGreenCalender()
        End Set
    End Property

    Private Sub SetupGreenCalender()
        Dim all As Action
        If Me.IsInitialized Then
            Dim act1 As Action = AddressOf SwitchTerm
            all = [Delegate].Combine(act1)
        Else
            Dim act1 As Action = AddressOf InitializeYearRange
            Dim act2 As Action = AddressOf GoYearRangeFoward
            Dim act3 As Action = AddressOf SetQuaterRange
            Dim act4 As Action = AddressOf SetMonthRange
            Dim act5 As Action = AddressOf SetWeekRange
            Dim act6 As Action = AddressOf SetYearIndex
            Dim act7 As Action = AddressOf SetQuaterIndex
            Dim act8 As Action = AddressOf SetYearRowHeader
            Dim act9 As Action = AddressOf SetQuaterRowHeader
            Dim act10 As Action = AddressOf SetYearColumnHeader
            Dim act11 As Action = AddressOf SetQuaterColumnHeader
            Dim act12 As Action = AddressOf SwitchTerm
            Dim act13 As Action = AddressOf SetYear
            all = [Delegate].Combine(act1, act2, act3, act4, act5, act6, act7, act8, act9, act10, act11, act12, act13)
            Me.IsInitialized = True
        End If
        Call all()
    End Sub

    Private Sub InitializeYearRange()
        If Me.YearRange.Count > 0 Then
            Dim latest As Date = Me.YearRange.Last.Day
            Dim term As Integer = (DateTime.Today - latest).TotalDays

            If term < 366 Then
                Exit Sub
            End If
        End If

        Me.YearRange = New ObservableCollection(Of CellData)
        Dim today As Date = DateTime.Today
        For i As Integer = 366 To 0 Step -1
            Me.YearRange.Add(New CellData With {.Day = today.AddDays(-i)})
        Next
    End Sub

    Private Sub GoYearRangeFoward()
        Dim latest As Date = Me.YearRange.Last.Day
        Dim term As Integer = (DateTime.Today - latest).TotalDays

        Dim i As Integer = term
        Do
            If i <= 0 Then
                Exit Do
            End If
            Me.YearRange.RemoveAt(0)
            Me.YearRange.Add(New CellData With {.Day = DateTime.Today.AddDays(-i + 1)})
            i -= 1
        Loop Until False
    End Sub

    Private Sub SetQuaterRange()
        Dim rng As IEnumerable(Of CellData)
        rng = Me.YearRange.Skip(Me.YearRange.Count - 90)
        For Each cell In rng
            Me.QuaterRange.Add(New CellData With {.Day = cell.Day, .Density = cell.Density})
        Next
    End Sub
    Private Sub SetMonthRange()
        Dim rng As IEnumerable(Of CellData)
        rng = Me.YearRange.Skip(Me.YearRange.Count - 31)
        For Each cell In rng
            Me.MonthRange.Add(New CellData With {.Day = cell.Day, .Density = cell.Density})
        Next
    End Sub
    Private Sub SetWeekRange()
        Dim rng As IEnumerable(Of CellData)
        rng = Me.YearRange.Skip(Me.YearRange.Count - 7)
        For Each cell In rng
            Me.WeekRange.Add(New CellData With {.Day = cell.Day, .Density = cell.Density})
        Next
    End Sub

    Private Sub SetYearIndex()
        Dim r As Integer = 0
        Dim c As Integer = 0
        For Each cell In Me.YearRange
            cell.Row = r
            cell.Column = c
            cell.RowSpan = 1
            cell.ColumnSpan = 1

            r += 1
            If r > 6 Then
                r = 0
                c += 1
            End If
        Next
    End Sub

    Private Sub SetQuaterIndex()
        Dim r As Integer = 0
        Dim c As Integer = 0
        For Each cell In Me.QuaterRange
            cell.Row = r
            cell.Column = c
            cell.RowSpan = 2
            cell.ColumnSpan = 2

            r += 2
            If r > 6 Then
                r = 0
                c += 2
            End If
        Next
    End Sub

    Private Sub SetYearRowHeader()
        Dim term As Integer = -1
        Dim premon As Integer = 0
        Dim col As Integer = 0

        Do
            Dim rhd As New RowHeaderData
            rhd.Column = col

            Dim mon As Double = 0.0
            Dim topmon As Integer = Me.YearRange(term + 1).Day.Month
            mon = topmon

            Dim nowmon As Integer = 0
            For j As Integer = 2 To 7
                Dim idx As Integer = term + j
                If idx <= (Me.YearRange.Count - 1) Then
                    nowmon = Me.YearRange(idx).Day.Month
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

            rhd.Header = ConvertDoubleToMonth(mon)
            rhd.ColumnSpan = 2
            Me.YearRowHeaders.Add(rhd)

            col += 1
            term += 7
            If term >= (Me.YearRange.Count - 1) Then
                Exit Do
            End If
        Loop Until False
    End Sub

    Private Sub SetQuaterRowHeader()
        Dim term As Integer = -1
        Dim preqtr As Double = 0.0
        Dim col As Integer = 0

        Do
            Dim rhd As New RowHeaderData
            rhd.Column = col

            Dim qtr As Double = 0.0
            Dim topday As Integer = Me.QuaterRange(term + 1).Day.Day
            Dim topqtr As Double = Me.QuaterRange(term + 1).Day.Month
            If topday <= 7 Then
                topqtr += 0.2
            ElseIf topday <= 14 Then
                topqtr += 0.4
            ElseIf topday <= 21 Then
                topqtr += 0.6
            ElseIf topday < 28 Then
                topqtr += 0.8
            Else
            End If
            qtr = topqtr

            Dim nowqtr As Double = 0.0
            For j As Integer = 2 To 4
                Dim idx As Integer = term + j
                If idx <= (Me.QuaterRange.Count - 1) Then
                    Dim nowday As Integer = Me.QuaterRange(idx).Day.Day
                    nowqtr = Me.QuaterRange(idx).Day.Month
                    If nowday <= 7 Then
                        nowqtr += 0.2
                    ElseIf nowday <= 14 Then
                        nowqtr += 0.4
                    ElseIf nowday <= 21 Then
                        nowqtr += 0.6
                    Else
                        nowqtr += 0.8
                    End If

                    If qtr <> nowqtr Then
                        qtr += 0.1
                        Exit For
                    End If
                End If
            Next

            If nowqtr <> topqtr Then
                preqtr = nowqtr
            Else
                If qtr = preqtr Then
                    qtr = 0.0
                Else
                    preqtr = nowqtr
                End If
            End If

            rhd.Header = ConvertDoubleToQuater(qtr)
            rhd.ColumnSpan = 3
            Me.QuaterRowHeaders.Add(rhd)

            col += 2
            term += 4
            If term >= (Me.QuaterRange.Count - 1) Then
                Exit Do
            End If
        Loop Until False
    End Sub

    Private Sub SetYearColumnHeader()
        Me.YearColumnHeaders.Add(New ColumnHeaderData With {
            .Row = 0,
            .Header = Strings.Left(Me.YearRange(0).Day.DayOfWeek.ToString, 3)
        })
        Me.YearColumnHeaders.Add(New ColumnHeaderData With {
            .Row = 3,
            .Header = Strings.Left(Me.YearRange(3).Day.DayOfWeek.ToString, 3)
        })
        Me.YearColumnHeaders.Add(New ColumnHeaderData With {
            .Row = 6,
            .Header = Strings.Left(Me.YearRange(6).Day.DayOfWeek.ToString, 3)
        })
    End Sub

    Private Sub SetQuaterColumnHeader()
    End Sub

    Public Sub SwitchTerm()
        Dim rowhead As ObservableCollection(Of RowHeaderData)
        Dim colhead As ObservableCollection(Of ColumnHeaderData)
        Dim cellrange As ObservableCollection(Of CellData)
        Select Case Me.TermIndex
            Case Me.TermTypes.IndexOf("Year")
                rowhead = Me.YearRowHeaders
                colhead = Me.YearColumnHeaders
                cellrange = Me.YearRange
            Case Me.TermTypes.IndexOf("Quater")
                rowhead = Me.QuaterRowHeaders
                colhead = Me.QuaterColumnHeaders
                cellrange = Me.QuaterRange
            Case Me.TermTypes.IndexOf("Month")
            Case Me.TermTypes.IndexOf("Week")
        End Select
        Me.RowHeaders = rowhead
        Me.ColumnHeaders = colhead
        Me.Range = cellrange
    End Sub


    Private Sub SetYear()
        Dim tyear As Integer = Me.YearRange.First.Day.Year.ToString
        Dim nyear As Integer = Me.YearRange.Last.Day.Year.ToString
        Me.YearString = $"({tyear} ~ {nyear})"
    End Sub

    Private Function ConvertDoubleToMonth(ByVal dbl As Double) As String
        Select Case dbl
            Case 0.0 : Return vbNullString
            Case 0.5 : Return "-/Jan."
            Case 1.0 : Return "Jan."
            Case 1.5 : Return "-/Feb."
            Case 2.0 : Return "Feb."
            Case 2.5 : Return "-/Mar."
            Case 3.0 : Return "Mar."
            Case 3.5 : Return "-/Apr."
            Case 4.0 : Return "Apr."
            Case 4.5 : Return "-/May"
            Case 5.0 : Return "May"
            Case 5.5 : Return "-/Jun."
            Case 6.0 : Return "Jun."
            Case 6.5 : Return "-/Jul."
            Case 7.0 : Return "Jul."
            Case 7.5 : Return "-/Aug."
            Case 8.0 : Return "Aug."
            Case 8.5 : Return "-/Sep."
            Case 9.0 : Return "Sep."
            Case 9.5 : Return "-/Oct."
            Case 10.0 : Return "Oct."
            Case 10.5 : Return "-/Nov."
            Case 11.0 : Return "Nov."
            Case 11.5 : Return "-/Dec."
            Case 12.0 : Return "Dec."
            Case 12.5 : Return "-/Jan."
            Case Else : Return vbNullString
        End Select
    End Function

    Private Function ConvertDoubleToQuater(ByVal dbl As Double) As String
        Select Case dbl
            Case 0.0 : Return vbNullString
            Case 0.2 : Return "Dec. 1st"
            Case 0.3 : Return "-/Dec. 2nd"
            Case 0.4 : Return "Dec. 2nd"
            Case 0.5 : Return "-/Dec. 3rd"
            Case 0.6 : Return "Dec. 3rd"
            Case 0.7 : Return "-/Dec. 4th"
            Case 0.8 : Return "Dec. 4th"
            Case 0.9 : Return "-/Jan. 1st"
            Case 1.2 : Return "Jan. 1st"
            Case 1.3 : Return "-/Jan. 2nd"
            Case 1.4 : Return "Jan. 2nd"
            Case 1.5 : Return "-/Jan. 3rd"
            Case 1.6 : Return "Jan. 3rd"
            Case 1.7 : Return "-/Jan. 4th"
            Case 1.8 : Return "Jan. 4th"
            Case 1.9 : Return "-/Feb. 1st"
            Case 2.2 : Return "Feb. 1st"
            Case 2.3 : Return "-/Feb. 2nd"
            Case 2.4 : Return "Feb. 2nd"
            Case 2.5 : Return "-/Feb. 3rd"
            Case 2.6 : Return "Feb. 3rd"
            Case 2.7 : Return "-/Feb. 4th"
            Case 2.8 : Return "Feb. 4th"
            Case 2.9 : Return "-/Mar. 1st"
            Case 3.2 : Return "Mar. 1st"
            Case 3.3 : Return "-/Mar. 2nd"
            Case 3.4 : Return "Mar. 2nd"
            Case 3.5 : Return "-/Mar. 3rd"
            Case 3.6 : Return "Mar. 3rd"
            Case 3.7 : Return "-/Mar. 4th"
            Case 3.8 : Return "Mar. 4th"
            Case 3.9 : Return "-/Apr. 1st"
            Case 4.2 : Return "Apr. 1st"
            Case 4.3 : Return "-/Apr. 2nd"
            Case 4.4 : Return "Apr. 2nd"
            Case 4.5 : Return "-/Apr. 3rd"
            Case 4.6 : Return "Apr. 3rd"
            Case 4.7 : Return "-/Apr. 4th"
            Case 4.8 : Return "Apr. 4th"
            Case 4.9 : Return "-/May 1st"
            Case 5.2 : Return "May 1st"
            Case 5.3 : Return "-/May 2nd"
            Case 5.4 : Return "May 2nd"
            Case 5.5 : Return "-/May 3rd"
            Case 5.6 : Return "May 3rd"
            Case 5.7 : Return "-/May 4th"
            Case 5.8 : Return "May 4th"
            Case 5.9 : Return "-/Jun. 1st"
            Case 6.2 : Return "Jun. 1st"
            Case 6.3 : Return "-/Jun. 2nd"
            Case 6.4 : Return "Jun. 2nd"
            Case 6.5 : Return "-/Jun. 3rd"
            Case 6.6 : Return "Jun. 3rd"
            Case 6.7 : Return "-/Jun. 4th"
            Case 6.8 : Return "Jun. 4th"
            Case 6.9 : Return "-/Jul. 1st"
            Case 7.2 : Return "Jul. 1st"
            Case 7.3 : Return "-/Jul. 2nd"
            Case 7.4 : Return "Jul. 2nd"
            Case 7.5 : Return "-/Jul. 3rd"
            Case 7.6 : Return "Jul. 3rd"
            Case 7.7 : Return "-/Jul. 4th"
            Case 7.8 : Return "Jul. 4th"
            Case 7.9 : Return "-/Aug. 1st"
            Case 8.2 : Return "Aug. 1st"
            Case 8.3 : Return "-/Aug. 2nd"
            Case 8.4 : Return "Aug. 2nd"
            Case 8.5 : Return "-/Aug. 3rd"
            Case 8.6 : Return "Aug. 3rd"
            Case 8.7 : Return "-/Aug. 4th"
            Case 8.8 : Return "Aug. 4th"
            Case 8.9 : Return "-/Sep. 1st"
            Case 9.2 : Return "Sep. 1st"
            Case 9.3 : Return "-/Sep. 2nd"
            Case 9.4 : Return "Sep. 2nd"
            Case 9.5 : Return "-/Sep. 3rd"
            Case 9.6 : Return "Sep. 3rd"
            Case 9.7 : Return "-/Sep. 4th"
            Case 9.8 : Return "Sep. 4th"
            Case 9.9 : Return "-/Oct. 1st"
            Case 10.2 : Return "Oct. 1st"
            Case 10.3 : Return "-/Oct. 2nd"
            Case 10.4 : Return "Oct. 2nd"
            Case 10.5 : Return "-/Oct. 3rd"
            Case 10.6 : Return "Oct. 3rd"
            Case 10.7 : Return "-/Oct. 4th"
            Case 10.8 : Return "Oct. 4th"
            Case 10.9 : Return "-/Nov. 1st"
            Case 11.2 : Return "Nov. 1st"
            Case 11.3 : Return "-/Nov. 2nd"
            Case 11.4 : Return "Nov. 2nd"
            Case 11.5 : Return "-/Nov. 3rd"
            Case 11.6 : Return "Nov. 3rd"
            Case 11.7 : Return "-/Nov. 4th"
            Case 11.8 : Return "Nov. 4th"
            Case 11.9 : Return "-/Dec. 1st"
            Case 12.2 : Return "Dec. 1st"
            Case 12.3 : Return "-/Dec. 2nd"
            Case 12.4 : Return "Dec. 2nd"
            Case 12.5 : Return "-/Dec. 3rd"
            Case 12.6 : Return "Dec. 3rd"
            Case 12.7 : Return "-/Dec. 4th"
            Case 12.8 : Return "Dec. 4th"
            Case 12.9 : Return "-/Jun. 1st"
            Case Else : Return vbNullString
        End Select
    End Function
End Class
