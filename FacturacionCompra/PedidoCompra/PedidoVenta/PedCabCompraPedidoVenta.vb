Public Class PedCabCompraPedidoVenta
    Inherits PedCabCompra

    Public IDBaseDatosSecundaria As Guid
    Public Multiempresa As Boolean
    Public Trazabilidad As Boolean
    Public EntregaProveedor As Boolean
    Public IDCliente As String
    Public DatosOrigen As DataTable

    Public Overrides Function FieldNOrigen() As String
        Return "NPedido"
    End Function

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return "IDPedido"
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        MyBase.Origen = enumOrigenPedidoCompra.PedidoVenta
        Trazabilidad = True
        Multiempresa = Nz(oRow("EmpresaGrupo"), False)
        If Multiempresa Then IDBaseDatosSecundaria = oRow("BaseDatos")
        EntregaProveedor = Nz(oRow("EntregaProveedor"), False)
        IDCliente = oRow("IDCliente")
    End Sub

End Class
