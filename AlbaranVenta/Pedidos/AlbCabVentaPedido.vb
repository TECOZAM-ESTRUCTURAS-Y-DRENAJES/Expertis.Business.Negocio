Public Class AlbCabVentaPedido
    Inherits AlbCabVenta

    Public Overrides Function FieldNOrigen() As String
        Return New String("NPedido")
    End Function

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return New String("IDPedido")
    End Function

    Public IDAlmacenDeposito As String
    Public EDI As Boolean
    Public Muelle As String
    Public PuntoDescarga As String
    Public Agrupacion As enummcAgrupAlbaran
    Public Intercambio As Boolean
    Public IDModoTransporte As String
    Public PedidoCliente As String
    Public Responsable As String
    Public IDClienteDistribuidor As String


    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        MyBase.Origen = enumOrigenAlbaranVenta.Pedido

        If oRow.Table.Columns.Contains("IDAlmacenDireccion") AndAlso Length(oRow("IDAlmacenDireccion")) > 0 Then IDAlmacenDeposito = oRow("IDAlmacenDireccion")
        If Length(oRow("IDFormaEnvio")) > 0 Then IDFormaEnvio = oRow("IDFormaEnvio")
        If Length(oRow("IDCondicionEnvio")) > 0 Then IDCondicionEnvio = oRow("IDCondicionEnvio")
        If Length(oRow("IDDireccionEnvio")) > 0 Then IdDireccion = oRow("IDDireccionEnvio")
        If Length(oRow("IDModoTransporte")) > 0 Then IDModoTransporte = oRow("IDModoTransporte")
        If Length(oRow("PedidoCliente")) > 0 Then PedidoCliente = oRow("PedidoCliente")

        If Length(oRow("IDFormaEnvio")) > 0 Then IDFormaEnvio = oRow("IDFormaEnvio")

        If Length(oRow("IDBancoPropio")) > 0 Then IDBancoPropio = oRow("IDBancoPropio")
        If Length(oRow("Responsable")) > 0 Then Responsable = oRow("Responsable")
        Dto = Nz(oRow("DtoPedido"), 0)
        EDI = oRow("EDI")
        Intercambio = oRow("Intercambio")
        Muelle = oRow("Muelle") & String.Empty
        PuntoDescarga = oRow("PuntoDescarga") & String.Empty
        If oRow.Table.Columns.Contains("TextoComercial") AndAlso Length(oRow("TextoComercial")) > 0 Then Me.ObsComerciales = oRow("TextoComercial")
        If oRow.Table.Columns.Contains("IDClienteDistribuidor") AndAlso Length(oRow("IDClienteDistribuidor")) > 0 Then IDClienteDistribuidor = oRow("IDClienteDistribuidor")

    End Sub

End Class
   