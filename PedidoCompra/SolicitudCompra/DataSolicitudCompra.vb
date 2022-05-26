<Serializable()> _
Public Class DataSolicitudCompra
    Public IDLineaSolicitud As Integer
    Public QSolicitar As Double

    Public Sub New(ByVal IDLineaSolicitud As Integer, ByVal QSolicitar As Double)
        Me.IDLineaSolicitud = IDLineaSolicitud
        Me.QSolicitar = QSolicitar
    End Sub
End Class
