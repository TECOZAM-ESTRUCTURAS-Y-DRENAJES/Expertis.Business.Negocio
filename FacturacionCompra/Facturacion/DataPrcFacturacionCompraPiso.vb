<Serializable()> _
Public Class DataPrcFacturacionCompraPiso

    Public IDPisoPago() As Integer
    Public IDContador As String
    Public DteFechaFactura As Date
    Public SuFactura As String
    Public IDPisoPagos As String

    Public Sub New(ByVal IDPisoPago() As Integer, ByVal IDContador As String, ByVal DteFechaFactura As Date, ByVal SuFactura As String, ByVal IDPisoPagos As String)
        Me.IDPisoPago = IDPisoPago
        Me.IDContador = IDContador
        Me.DteFechaFactura = DteFechaFactura
        Me.SuFactura = SuFactura
        Me.IDPisoPagos = IDPisoPagos
    End Sub
End Class