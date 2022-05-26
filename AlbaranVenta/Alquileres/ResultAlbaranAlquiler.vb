<Serializable()> _
Public Class ResultAlbaranAlquiler
    Inherits AlbaranLogProcess
    Public PropuestaAlbaranes As DataTable       '//Para la Propuesta

    Public Sub New()
        CreateData = New LogProcess
    End Sub

    Public Sub New(ByVal n As Integer, ByVal id As Integer, ByVal LogProc As LogProcess)
        CreateData = LogProc
    End Sub

    Public Sub New(ByVal dtAlbaranes As DataTable)
        PropuestaAlbaranes = dtAlbaranes
        CreateData = New LogProcess
    End Sub
End Class
