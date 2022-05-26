Public Class FraCabCompraAlbaran
    Inherits FraCabCompra

    Public IDAlbaran As Integer
    Public NAlbaran As String
    Public Texto As String

    Public Lineas(-1) As FraLinCompraAlbaran

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(New DataRowPropertyAccessor(oRow))

        IDAlbaran = oRow("IDAlbaran")
        NAlbaran = oRow("NAlbaran")
        'Fecha = oRow("FechaAlbaran")
        Texto = oRow("Texto") & String.Empty
    End Sub

    Public Sub Add(ByVal lin As FraLinCompraAlbaran)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub
End Class
