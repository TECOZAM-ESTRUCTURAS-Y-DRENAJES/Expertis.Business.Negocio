<Serializable()> _
Public Class DataResultadoMultiempresaPC

    Private htElementosGenerados As New System.Collections.Generic.Dictionary(Of Integer, GeneracionPedidosCompraInfo)

    Public Sub Add(ByVal data As GeneracionPedidosCompraInfo)
        htElementosGenerados(data.IDPedidoCompra) = data
    End Sub

    Public Function Items() As System.Collections.Generic.Dictionary(Of Integer, GeneracionPedidosCompraInfo)
        Return htElementosGenerados
    End Function

    Public Function Item(ByVal IDPedido As Integer) As GeneracionPedidosCompraInfo
        Return htElementosGenerados(IDPedido)
    End Function

    'Public Function ExisteElemento(ByVal IDPedido As Integer) As Boolean
    '    Return htElementosGenerados.ContainsKey(IDPedido)
    'End Function

End Class
