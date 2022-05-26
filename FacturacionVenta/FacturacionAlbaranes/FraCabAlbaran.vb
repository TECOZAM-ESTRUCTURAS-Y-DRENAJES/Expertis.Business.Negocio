Public Class FraCabAlbaran
    Inherits FraCab
    Public IDAlbaran As Integer
    Public NAlbaran As String
    Public IDTPV As String
    Public AgrupFactura As enummcAgrupFactura

    Public Lineas(-1) As FraLinAlbaran

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(New DataRowPropertyAccessor(oRow))
        IDAlbaran = oRow("IDAlbaran")
        NAlbaran = oRow("NAlbaran")
        Fecha = oRow("FechaAlbaran")
        Dto = oRow("DtoAlbaran")
        If Length(oRow("IDTPV")) > 0 Then IDTPV = oRow("IDTPV")
        ObsComerciales = oRow("Texto") & String.Empty
        AgrupFactura = oRow("AgrupFactura")
    End Sub

    Public Sub Add(ByVal lin As FraLinAlbaran)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub
End Class
