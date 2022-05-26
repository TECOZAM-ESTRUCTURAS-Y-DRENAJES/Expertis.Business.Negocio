Public Class PedCabVentaPedidoCompra
    Inherits PedCab

    Public Multiempresa As Boolean
    Public Trazabilidad As Boolean
    Public EntregaProveedor As Boolean
    Public DatosOrigen As DataTable '//Tienen que viajar los datos por que no podemos recuperar los PC desde la BBDD Secundaria
    Public PedidoCliente As String
    Public Texto As String
    Public IDPedido As Integer
    Public NPedido As String
    Public IDDireccionEnvio As Integer

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        MyBase.Origen = enumOrigenPedido.PedidoCompra

        Trazabilidad = True
        Multiempresa = Nz(oRow("EmpresaGrupo"), False)
        EntregaProveedor = Nz(oRow("EntregaProveedor"), False)
        'IDCliente = oRow("IDCliente")
        IDPedido = oRow("IDPedido")  '//IDPCPrincipal
        NPedido = oRow("NPedido")    '//NPCPrincipal
        Fecha = Today
        If oRow.Table.Columns.Contains("PedidoCliente") AndAlso Length(oRow("PedidoCliente")) > 0 Then PedidoCliente = oRow("PedidoCliente")
        If Length(oRow("Texto")) > 0 Then Texto = oRow("Texto")
    End Sub

End Class
