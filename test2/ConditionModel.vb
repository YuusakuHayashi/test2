Imports System.Collections.ObjectModel

Public Class ConditionModel
    Private _FieldName As String
    Public Property FieldName As String
        Get
            Return Me._FieldName
        End Get
        Set(value As String)
            Me._FieldName = value
        End Set
    End Property

    Private _FieldValue As Object
    Public Property FieldValue As Object
        Get
            Return Me._FieldValue
        End Get
        Set(value As Object)
            Me._FieldValue = value
        End Set
    End Property

    Private _Conditions As ObservableCollection(Of ConditionModel)
    Public Property Conditions As ObservableCollection(Of ConditionModel)
        Get
            Return Me._Conditions
        End Get
        Set(value As ObservableCollection(Of ConditionModel))
            Me._Conditions = value
        End Set
    End Property
End Class
