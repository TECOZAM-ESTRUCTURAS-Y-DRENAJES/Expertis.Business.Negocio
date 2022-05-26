Public Class FraLinCompraAlbaran

    Public IDAlbaran As Integer
    Public IDLineaAlbaran As Integer
    Public QaFacturar As Double
    Public QIntAFacturar As Double

    Public Sub New(ByVal oRow As DataRow)
        IDAlbaran = oRow("IDAlbaran")
        IDLineaAlbaran = oRow("IDLineaAlbaran")
        QaFacturar = Double.NaN
        QIntAFacturar = Double.NaN
    End Sub

End Class
