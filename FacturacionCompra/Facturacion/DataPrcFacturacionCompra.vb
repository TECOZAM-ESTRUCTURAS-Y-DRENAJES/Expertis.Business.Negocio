<Serializable()> _
Public Class DataPrcFacturacionCompra

    Public IDAlbaranes() As Integer
    Public IDContador As String
    Public DteFechaFactura As Date
    Public SuFactura As String
    Public Sub New(ByVal IDAlbaranes() As Integer, ByVal IDContador As String, ByVal DteFechaFactura As Date, ByVal SuFactura As String)
        Me.IDAlbaranes = IDAlbaranes
        Me.IDContador = IDContador
        Me.DteFechaFactura = DteFechaFactura
        Me.SuFactura = SuFactura
    End Sub
End Class
