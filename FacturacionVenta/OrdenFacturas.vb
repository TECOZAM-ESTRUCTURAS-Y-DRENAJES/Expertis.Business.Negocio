Public Class OrdenFacturas
    Implements IComparer
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        Dim reslt As Integer = DirectCast(x, FraCabAlbaran).Fecha.CompareTo(DirectCast(y, FraCabAlbaran).Fecha)
        If reslt = 0 Then reslt = DirectCast(x, FraCabAlbaran).IDAlbaran.CompareTo(DirectCast(y, FraCabAlbaran).IDAlbaran)
        Return reslt
    End Function
End Class
