Public Class FraCabCompraPiso
    Inherits FraCabCompra

    Public IDPisoPago As Integer
    Public NAlbaran As String
    Public Texto As String

    Public Lineas(-1) As FraLinCompraPiso

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(New DataRowPropertyAccessor(oRow))

        IDPisoPago = oRow("IDPisoPago")
        'NAlbaran = oRow("NAlbaran")
        'Fecha = oRow("FechaAlbaran")
        'Texto = oRow("Texto") & String.Empty
    End Sub

    Public Sub Add(ByVal lin As FraLinCompraPiso)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub
End Class
