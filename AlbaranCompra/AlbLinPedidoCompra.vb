Public Class AlbLinPedidoCompra

    Public IDPedido As Integer
    Public IDLineaPedido As Integer
    Public FechaEntregaModificado As Date
    Public QaRecibir As Double
    Public Cantidad As Double
    Public Cantidad2 As Double?
    Public Lotes As DataTable
    Public Series As DataTable

    Public Sub New(ByVal oRow As DataRow)
        IDPedido = oRow("IDPedido")
        IDLineaPedido = oRow("IDLineaPedido")
        FechaEntregaModificado = Nz(oRow("FechaEntregaModificado"), cnMinDate)
        QaRecibir = Double.NaN
        Cantidad = Double.NaN
    End Sub
End Class

