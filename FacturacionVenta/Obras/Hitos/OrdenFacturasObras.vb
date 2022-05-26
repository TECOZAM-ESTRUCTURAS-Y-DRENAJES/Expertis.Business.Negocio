Public Class OrdenFacturasObras
    Implements IComparer
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        Dim reslt As Integer = DirectCast(x, FraCabObra).Fecha.CompareTo(DirectCast(y, FraCabObra).Fecha)
        If reslt = 0 Then reslt = DirectCast(x, FraCabObra).IDObra.CompareTo(DirectCast(y, FraCabObra).IDObra)
        Return reslt
    End Function
End Class