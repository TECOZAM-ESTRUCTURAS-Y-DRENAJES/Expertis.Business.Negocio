Public Class AlbCabVentaMaquinaria
    Inherits AlbCabVentaObras

    Public IDTipoAlbaran As String

    Public dtActivos As DataTable

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        Me.Origen = enumOrigenAlbaranVenta.Alquiler
        Me.Fecha = oRow("FechaAlbaran")
    End Sub

End Class
