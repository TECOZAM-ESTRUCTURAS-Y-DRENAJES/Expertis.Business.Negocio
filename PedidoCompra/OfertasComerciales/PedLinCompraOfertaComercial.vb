Public Class PedLinCompraOfertaComercial
    Inherits PedLinCompra

    'Public IDMarca As String
    'Public IDArticulo As String
    'Public IDAlmacen As String

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return "IDLineaOfertaDetalle"
    End Function

    
    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        Me.Cantidad = Nz(oRow("QEstimadaConsumo"), 0)
    End Sub

End Class
