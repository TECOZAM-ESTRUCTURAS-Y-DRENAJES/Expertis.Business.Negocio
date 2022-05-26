Public Class AlbCabVentaAlquiler
    Inherits AlbCabVentaObras

    Public Overrides Function FieldNOrigen() As String
        Return New String("NObra")
    End Function

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return New String("IDObra")
    End Function

    Public IDAlmacen As String
    Public IDAlmacenDeposito As String
    Public IDAlmacenTransferencia As String
    Public Deposito As Boolean = False

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        Me.Origen = enumOrigenAlbaranVenta.Alquiler
        Me.Fecha = oRow("FechaAlquiler")
        Me.IDAlmacen = oRow("IDAlmacen") & String.Empty
        If oRow.Table.Columns.Contains("IDAlmacenDeposito") Then Me.IDAlmacenDeposito = oRow("IDAlmacenDeposito") & String.Empty
        If oRow.Table.Columns.Contains("IDAlmacenTransferencia") Then Me.IDAlmacenTransferencia = oRow("IDAlmacenTransferencia") & String.Empty
        If oRow.Table.Columns.Contains("Deposito") Then Me.Deposito = Nz(oRow("Deposito"), False)
        'If oRow.Table.Columns.Contains("NumeroPedido") Then If Length(oRow("NumeroPedido")) > 0 Then Me.PedidoCliente = oRow("NumeroPedido")
        If oRow.Table.Columns.Contains("PedidoCliente") Then Me.PedidoCliente = oRow("PedidoCliente") & String.Empty
        Me.Texto = oRow("TextoPublico") & String.Empty
        If oRow.Table.Columns.Contains("IDAlbaran") Then Me.IDOrigen = oRow("IDAlbaran")
    End Sub

End Class
