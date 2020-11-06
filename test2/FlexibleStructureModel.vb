Public Class FlexibleStructureModel
    Private _MainStructure As FlexibleStructureModel
    Public Property MainStructure As FlexibleStructureModel
        Get
            Return Me._MainStructure
        End Get
        Set(value As FlexibleStructureModel)
            Me._MainStructure = value
        End Set
    End Property

    Private _RightStructure As FlexibleStructureModel
    Public Property RightStructure As FlexibleStructureModel
        Get
            Return Me._RightStructure
        End Get
        Set(value As FlexibleStructureModel)
            Me._RightStructure = value
        End Set
    End Property

    Private _BottomStructure As FlexibleStructureModel
    Public Property BottomStructure As FlexibleStructureModel
        Get
            Return Me._BottomStructure
        End Get
        Set(value As FlexibleStructureModel)
            Me._BottomStructure = value
        End Set
    End Property

    Private _Name As String
    Public Property [Name] As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
        End Set
    End Property

    Private _ModelName As String
    Public Property ModelName As String
        Get
            Return Me._ModelName
        End Get
        Set(value As String)
            Me._ModelName = value
        End Set
    End Property
End Class
