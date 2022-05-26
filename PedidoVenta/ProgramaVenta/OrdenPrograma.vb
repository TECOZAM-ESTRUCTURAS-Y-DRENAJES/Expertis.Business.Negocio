Public Class OrdenPrograma
    Implements IComparer
    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        Dim reslt As Integer = DirectCast(x, PedCabPrograma).IDCliente.CompareTo(DirectCast(y, PedCabPrograma).IDCliente)
        If reslt = 0 Then reslt = DirectCast(x, PedCabPrograma).IDMoneda.CompareTo(DirectCast(y, PedCabPrograma).IDMoneda)
        If reslt = 0 Then reslt = DirectCast(x, PedCabPrograma).IDPrograma.CompareTo(DirectCast(y, PedCabPrograma).IDPrograma)
        Return reslt
    End Function
End Class
