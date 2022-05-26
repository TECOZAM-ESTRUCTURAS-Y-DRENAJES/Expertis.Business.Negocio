<Serializable()> _
Public Class DataPrcFacturacionCompraLeasing

    Public IDPagos() As Integer
    Public IDContador As String
    'Public DteFechaFactura As Date
    'Public SuFactura As String
    Public Sub New(ByVal IDPagos() As Integer, ByVal IDContador As String) ', ByVal DteFechaFactura As Date, ByVal SuFactura As String)
        Me.IDPagos = IDPagos
        Me.IDContador = IDContador
        'Me.DteFechaFactura = DteFechaFactura
        'Me.SuFactura = SuFactura
    End Sub

End Class
