Public Class DataPrcCrearPedidoVentaEnBDSecundaria
    Public Doc As DocumentoPedidoCompra
    Public IDBaseDatosPrincipal As Guid
    Public IDBaseDatosSecundaria As Guid
    Public IDDireccion As Integer
    Public IDPedidoVenta As Integer?         '//IDPedido Venta Origen. Si viene de PC manual este vendrá vacío
    Public NPedidoVenta As String            '//NPedido Venta Origen
    Public IDCliente As String

    Public Sub New(ByVal IDBaseDatosPrincipal As Guid, ByVal Doc As DocumentoPedidoCompra)
        Me.IDBaseDatosPrincipal = IDBaseDatosPrincipal
        'Me.IDBaseDatosSecundaria = IDBaseDatosSecundaria
        Me.IDDireccion = Doc.HeaderRow("IDDireccion")
        If Not Doc.Cabecera Is Nothing Then
            Me.IDPedidoVenta = Doc.Cabecera.IDOrigen
            Me.NPedidoVenta = Doc.Cabecera.NOrigen
        End If
        Me.Doc = Doc
    End Sub

End Class
