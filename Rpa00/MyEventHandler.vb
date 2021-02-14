Public Class MyEventHandler
    ' GuiでViewController.DataのPropertyChangedを購読したくても、
    ' dllから作成したObject型データなので、不可
    ' よって、独自イベントを発行して、それを購読している
    Private Shared _Instance As MyEventHandler = New MyEventHandler
    Public Shared ReadOnly Property Instance As MyEventHandler
        Get
            Return MyEventHandler._Instance
        End Get
    End Property

    Public Event CommandLogsChanged As EventHandler
    Public Sub RaiseCommandLogsChanged(ByVal logs As List(Of RpaCommand))
        RaiseEvent CommandLogsChanged(logs, EventArgs.Empty)
    End Sub
End Class
