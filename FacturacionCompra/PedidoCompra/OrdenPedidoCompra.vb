Public Class OrdenPedidoCompra
    Implements IComparer
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        Dim reslt As Integer = DirectCast(x, PedCabCompra).IDProveedor.CompareTo(DirectCast(y, PedCabCompra).IDProveedor)
        If reslt = 0 Then reslt = DirectCast(x, PedCabCompra).IDMoneda.CompareTo(DirectCast(y, PedCabCompra).IDMoneda)
        If reslt = 0 Then reslt = DirectCast(x, PedCabCompra).IDCondicionPago.CompareTo(DirectCast(y, PedCabCompra).IDCondicionPago)
        If reslt = 0 Then reslt = DirectCast(x, PedCabCompra).IDFormaPago.CompareTo(DirectCast(y, PedCabCompra).IDFormaPago)
        Return reslt
    End Function
End Class