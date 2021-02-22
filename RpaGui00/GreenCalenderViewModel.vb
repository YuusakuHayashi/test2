Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports Newtonsoft.Json

Public Class GreenCalenderViewModel : Inherits ViewModelBase(Of GreenCalenderViewModel)
    Public Class GreenDay : Implements INotifyPropertyChanged

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
        Protected Sub RaisePropertyChanged(ByVal PropertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
        End Sub

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

        Public ReadOnly Property Week As String
            Get
                Return Me.Day.DayOfWeek.ToString
            End Get
        End Property
    End Class

    Public Class ColumnHeaderData : Implements INotifyPropertyChanged
        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
        Protected Sub RaisePropertyChanged(ByVal PropertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(PropertyName))
        End Sub

        Private _Name As String
        Public Property Name As String
            Get
                Return Me._Name
            End Get
            Set(value As String)
                Me._Name = value
                RaisePropertyChanged("Name")
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

    Private _GreenYear As ObservableCollection(Of GreenDay)
    Public Property GreenYear As ObservableCollection(Of GreenDay)
        Get
            If Me._GreenYear Is Nothing Then
                Me._GreenYear = New ObservableCollection(Of GreenDay)
            End If
            Return Me._GreenYear
        End Get
        Set(value As ObservableCollection(Of GreenDay))
            Me._GreenYear = value
            RaisePropertyChanged("GreenYear")
        End Set
    End Property

    Private _Months As ObservableCollection(Of Double)
    <JsonIgnore>
    Public Property Months As ObservableCollection(Of Double)
        Get
            If Me._Months Is Nothing Then
                Me._Months = New ObservableCollection(Of Double)
                For i As Integer = 0 To 50
                    Me._Months.Add(0.0)
                Next
            End If
            Return Me._Months
        End Get
        Set(value As ObservableCollection(Of Double))
            Me._Months = value
            RaisePropertyChanged("Months")
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
        Dim act3 As Action = AddressOf SetMonthRow
        Dim act4 As Action = AddressOf SetWeekColumn
        Dim act5 As Action = AddressOf SetYear
        Dim all As Action = [Delegate].Combine(act1, act2, act3, act4, act5)
        Call all()
    End Sub

    Private Sub InitializeCalender()
        If Me.GreenYear.Count > 0 Then
            Dim latest As Date = Me.GreenYear.Last.Day
            Dim term As Integer = (DateTime.Today - latest).TotalDays

            If term < 366 Then
                Exit Sub
            End If
        End If

        Me.GreenYear = New ObservableCollection(Of GreenDay)
        Dim today As Date = DateTime.Today
        For i As Integer = 366 To 0 Step -1
            Me.GreenYear.Add(New GreenDay With {.Day = today.AddDays(-i)})
        Next
    End Sub

    Private Sub GoCalenderFoward()
        Dim latest As Date = Me.GreenYear.Last.Day
        Dim term As Integer = (DateTime.Today - latest).TotalDays

        Dim i As Integer = term
        Do
            If i <= 0 Then
                Exit Do
            End If
            Me.GreenYear.RemoveAt(0)
            Me.GreenYear.Add(New GreenDay With {.Day = DateTime.Today.AddDays(-i + 1)})
            i -= 1
        Loop Until False
    End Sub

    Private Sub SetMonthRow()
        'Dim term As Integer = 367
        Dim term As Integer = -1
        Dim premon As Integer = 0
        Dim mons() As Double = Me.Months.ToArray

        For i As Integer = LBound(mons) To UBound(mons)
            Dim nowmon As Integer = 0
            Dim topmon As Integer = Me.GreenYear(term + 1).Day.Month
            mons(i) = topmon
            For j As Integer = 2 To 7
                nowmon = Me.GreenYear(term + j).Day.Month
                If mons(i) <> nowmon Then
                    mons(i) += 0.5
                    Exit For
                End If
            Next

            If mons(i) <> topmon Then
                premon = nowmon
            Else
                If mons(i) = premon Then
                    mons(i) = 0.0
                Else
                    premon = nowmon
                End If
            End If

            Me.Months(i) = mons(i)

            term += 7
            If term >= 366 Then
                Exit For
            End If
        Next
    End Sub

    Private Sub SetWeekColumn()
        Me.ColumnHeaders.Add(New ColumnHeaderData With {
            .Row = 1,
            .Name = Strings.Left(Me.GreenYear(0).Day.DayOfWeek.ToString, 3)
        })
        Me.ColumnHeaders.Add(New ColumnHeaderData With {
            .Row = 4,
            .Name = Strings.Left(Me.GreenYear(3).Day.DayOfWeek.ToString, 3)
        })
        Me.ColumnHeaders.Add(New ColumnHeaderData With {
            .Row = 7,
            .Name = Strings.Left(Me.GreenYear(6).Day.DayOfWeek.ToString, 3)
        })
    End Sub

    Private Sub SetYear()
        Dim tyear As Integer = Me.GreenYear.First.Day.Year.ToString
        Dim nyear As Integer = Me.GreenYear.Last.Day.Year.ToString
        Me.YearString = $"({tyear} ~ {nyear})"
    End Sub
End Class
