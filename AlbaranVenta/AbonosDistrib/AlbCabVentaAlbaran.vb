Public Class AlbCabVentaAlbaran
    Inherits AlbCabVenta

    Public Overrides Function FieldNOrigen() As String
        Return New String("NAlbaran")
    End Function

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return New String("IDAlbaran")
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        MyBase.Origen = enumOrigenAlbaranVenta.AlbaranDistrib

        'If oRow.Table.Columns.Contains("IDAlmacenDireccion") AndAlso Length(oRow("IDAlmacenDireccion")) > 0 Then IDAlmacenDeposito = oRow("IDAlmacenDireccion")
        If Length(oRow("IDFormaEnvio")) > 0 Then IDFormaEnvio = oRow("IDFormaEnvio")
        If Length(oRow("IDCondicionEnvio")) > 0 Then IDCondicionEnvio = oRow("IDCondicionEnvio")
        If Length(oRow("IDDireccion")) > 0 Then IdDireccion = oRow("IDDireccion")
        If Length(oRow("IDModoTransporte")) > 0 Then IDModoTransporte = oRow("IDModoTransporte")
        'If Length(oRow("PedidoCliente")) > 0 Then PedidoCliente = oRow("PedidoCliente")

        If Length(oRow("IDFormaEnvio")) > 0 Then IDFormaEnvio = oRow("IDFormaEnvio")

        If Length(oRow("IDBancoPropio")) > 0 Then IDBancoPropio = oRow("IDBancoPropio")
        'If Length(oRow("Responsable")) > 0 Then Responsable = oRow("Responsable")
        Dto = Nz(oRow("DtoAlbaran"), 0)
        EDI = oRow("EDI")
        'Intercambio = oRow("Intercambio")
        'Muelle = oRow("Muelle") & String.Empty
        'PuntoDescarga = oRow("PuntoDescarga") & String.Empty
        If oRow.Table.Columns.Contains("TextoComercial") AndAlso Length(oRow("TextoComercial")) > 0 Then Me.ObsComerciales = oRow("TextoComercial")
    End Sub


End Class
