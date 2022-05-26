Public Class OrdenFacturasCompra
    Implements IComparer
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        Dim reslt As Integer = DirectCast(x, FraCabCompra).Fecha.CompareTo(DirectCast(y, FraCabCompra).Fecha)
        If TypeOf x Is FraCabCompraAlbaran Then If reslt = 0 Then reslt = DirectCast(x, FraCabCompraAlbaran).IDAlbaran.CompareTo(DirectCast(y, FraCabCompraAlbaran).IDAlbaran)
        Return reslt
    End Function
End Class
