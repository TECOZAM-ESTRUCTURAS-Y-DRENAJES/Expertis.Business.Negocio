Public Class AlbLinVentaPedido
    Inherits AlbLinVenta

    Public PedidoCliente As String
    Public FechaEntregaModificado As Date?
    Public Cantidad2 As Double?
    'Public Lotes As DataTable
    'Public Series As DataTable

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return New String("IDLineaPedido")
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        If Length(oRow("PedidoCliente")) > 0 Then PedidoCliente = oRow("PedidoCliente")
        'If Length(oRow("FechaEntregaModificado")) > 0 Then FechaEntregaModificado = oRow("FechaEntregaModificado")
    End Sub

End Class

