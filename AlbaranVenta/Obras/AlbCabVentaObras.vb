Public Class AlbCabVentaObras
    Inherits AlbCabVenta

    Public PedidoCliente As String
    Public Texto As String

    Public Overrides Function FieldNOrigen() As String
        Return New String("NObra")
    End Function

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return New String("IDObra")
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        MyBase.Origen = enumOrigenAlbaranVenta.Obras
        If oRow.Table.Columns.Contains("IDDireccion") AndAlso Length(oRow("IDDireccion")) > 0 Then IdDireccion = oRow("IdDireccion").ToString
        If oRow.Table.Columns.Contains("IDCondicionEnvio") AndAlso Length(oRow("IDCondicionEnvio")) > 0 Then IDCondicionEnvio = oRow("IDCondicionEnvio")
        If oRow.Table.Columns.Contains("NumeroPedido") AndAlso Length(oRow("NumeroPedido")) > 0 Then PedidoCliente = oRow("NumeroPedido")
        If oRow.Table.Columns.Contains("Texto") Then Me.Texto = oRow("Texto") & String.Empty
    End Sub

End Class
