Public Class PedCabCompraOfertaCompra
    Inherits PedCabCompra

    Public IDOperario As String
    Public IDDiaPago As String

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return "IDOferta"
    End Function

    Public Overrides Function FieldNOrigen() As String
        Return String.Empty
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        MyBase.ViewName = "vctlConsOfertasCompraOfertasCompra"
        MyBase.Origen = enumOrigenPedidoCompra.OfertaCompra
        If Length(oRow("IDOperarioOF")) > 0 Then IDOperario = oRow("IDOperarioOf")
        If Length(oRow("IDDiaPagoOf")) > 0 Then IDDiaPago = oRow("IDDiaPagoOf")
    End Sub

End Class
