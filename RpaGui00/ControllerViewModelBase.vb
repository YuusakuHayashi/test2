Imports Newtonsoft.Json
Public Class ControllerViewModelBase(Of T As {New}) : Inherits ViewModelBase(Of T)
    Private _ViewController As ControllerViewModel
    <JsonIgnore>
    Public Property ViewController As ControllerViewModel
        Get
            Return Me._ViewController
        End Get
        Set(value As ControllerViewModel)
            Me._ViewController = value
            Call RaisePropertyChanged("ViewController")
        End Set
    End Property
End Class
