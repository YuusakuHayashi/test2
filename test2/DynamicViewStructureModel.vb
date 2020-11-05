Public Class DynamicViewStructureModel
    Private _MainContent As DynamicViewStructureModel
    Public Property MainContent As DynamicViewStructureModel
        Get
            Return Me._MainContent
        End Get
        Set(value As DynamicViewStructureModel)
            Me._MainContent = value
        End Set
    End Property

    Private _RightContent As DynamicViewStructureModel
    Public Property RightContent As DynamicViewStructureModel
        Get
            Return Me._RightContent
        End Get
        Set(value As DynamicViewStructureModel)
            Me._RightContent = value
        End Set
    End Property

    Private _BottomContent As DynamicViewStructureModel
    Public Property BottomContent As DynamicViewStructureModel
        Get
            Return Me._BottomContent
        End Get
        Set(value As DynamicViewStructureModel)
            Me._BottomContent = value
        End Set
    End Property
End Class
