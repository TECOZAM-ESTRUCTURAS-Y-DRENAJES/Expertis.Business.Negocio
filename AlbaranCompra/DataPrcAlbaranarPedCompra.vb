'//Clase q proporciona datos de entrada al proceso
<Serializable()> _
Public Class DataPrcAlbaranarPedCompra
    Public Pedidos() As CrearAlbaranCompraInfo
    Public IDContador As String
    Public FechaAlbaran As Date
    Public IDTipoCompra As String

    Public Sub New(ByVal PedidosInfo() As CrearAlbaranCompraInfo, Optional ByVal IDContador As String = Nothing, Optional ByVal FechaAlbaran As Date = cnMinDate, Optional ByVal IDTipoCompra As String = Nothing)
        Me.Pedidos = PedidosInfo
        Me.IDContador = IDContador
        Me.FechaAlbaran = FechaAlbaran
        Me.IDTipoCompra = IDTipoCompra
    End Sub

    Public Sub New(ByVal IDPedido As Integer, Optional ByVal IDTipoCompra As String = Nothing)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDPedido", FilterOperator.Equal, IDPedido))
        Dim dtCabPedido As DataTable = New PedidoCompraCabecera().SelOnPrimaryKey(IDPedido)
        Dim dtPedidos As DataTable = New PedidoCompraLinea().Filter(f)

        If Not dtPedidos Is Nothing AndAlso dtPedidos.Rows.Count > 0 Then
            ReDim Preserve Pedidos(-1)
            For Each drPedido As DataRow In dtPedidos.Rows
                If drPedido("Estado") = enumpclEstado.pclpedido Then
                    Dim dataInfo As New CrearAlbaranCompraInfo
                    dataInfo.IDLinea = drPedido("IDLineaPedido")
                    dataInfo.Cantidad = drPedido("QPedida") - drPedido("QServida")
                    dataInfo.CantidadUD = drPedido("QInterna")
                    If drPedido.Table.Columns.Contains("QInterna2") AndAlso Length(drPedido("QInterna2")) > 0 Then dataInfo.Cantidad2 = CDbl(drPedido("QInterna2"))
                    If Length(IDTipoCompra) > 0 Then
                        dataInfo.IDTipoCompra = IDTipoCompra
                    Else
                        dataInfo.IDTipoCompra = dtCabPedido.Rows(0)("IDTipoCompra")
                    End If

                    Me.IDTipoCompra = dataInfo.IDTipoCompra
                    ReDim Preserve Pedidos(UBound(Pedidos) + 1)
                    Pedidos(UBound(Pedidos)) = dataInfo
                End If
            Next

        End If
    End Sub

End Class
