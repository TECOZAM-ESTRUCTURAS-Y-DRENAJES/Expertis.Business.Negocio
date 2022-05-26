Public MustInherit Class AlbLinVenta

    Public MustOverride Function PrimaryKeyLinOrigen() As String

    Public IDLineaOrigen As Integer
    Public QaServir As Double
    Public Cantidad As Double
    Public Lotes As DataTable
    Public Series As DataTable
    Public Seguimiento As DataTable
    Public ArtCompatibles As DataArtCompatiblesExp


    Public Sub New(ByVal oRow As DataRow)
        If Length(PrimaryKeyLinOrigen) > 0 Then IDLineaOrigen = nz(oRow(PrimaryKeyLinOrigen), 0)
        QaServir = Double.NaN
    End Sub

End Class
