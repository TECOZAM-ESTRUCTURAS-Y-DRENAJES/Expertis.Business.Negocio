Public Class FraLinObraPromo

    Public IDLocalVencimiento As Integer
    Public DescVencimiento As String
    Public IDArticulo As String
    Public IDLocal As Integer
    Public Fecha As Date
    Public IDTipoIva As String
    Public Edificio As String
    Public Piso As String
    Public Letra As String
    Public DireccionObra As String
    Public NumeroGaraje As String
    Public Descripcion2 As String
    Public Descripcion3 As String
    Public Descripcion4 As String

    Public Sub New(ByVal oRow As DataRow)
        Me.IDLocalVencimiento = oRow("IDLocalVencimiento")
        Me.DescVencimiento = oRow("DescVencimiento")
        Me.IDArticulo = oRow("IDArticulo")
        Me.IDLocal = oRow("IDLocal")
        Me.Fecha = Nz(oRow("FechaVencimiento"), Date.Today)
        Me.IDTipoIva = oRow("IDTipoIva")
        Me.Edificio = oRow("Edificio") & String.Empty
        Me.Piso = oRow("Piso") & String.Empty
        Me.Letra = oRow("Letra") & String.Empty
        Me.NumeroGaraje = oRow("NumeroGarage") & String.Empty
        'Construcción de la dirección completa de la Obra
        Me.DireccionObra = oRow("Direccion") & String.Empty
        If Length(oRow("Poblacion")) > 0 Then Me.DireccionObra = DireccionObra & " de " & oRow("Poblacion")
        Me.Descripcion2 = oRow("Descripcion2") & String.Empty
        Me.Descripcion3 = oRow("Descripcion3") & String.Empty
        Me.Descripcion4 = oRow("Descripcion4") & String.Empty
    End Sub

End Class