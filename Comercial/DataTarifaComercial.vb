<Serializable()> _
Public Class DataTarifaComercial
    Inherits DataTarifa

    Public IDUDVenta As String      ' Dato de Salida
    Public PVP As Double

    Public IDTarifa As String
    Public IDPromocion As String
    Public IDPromocionLinea As Integer?
    Public IDLineaOfertaDetalle As Integer?

    Public SeguimientoDtos As String

    Public PrecioCosteA As Double
    Public PrecioCosteB As Double
End Class
