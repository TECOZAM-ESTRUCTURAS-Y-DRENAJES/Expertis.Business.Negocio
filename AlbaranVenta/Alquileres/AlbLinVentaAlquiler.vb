Public Class AlbLinVentaAlquiler
    Inherits AlbLinVentaObras

    Public HoraAlquiler As Date
    Public IDAlmacenDeposito As String
    Public IDAlmacen As String
    Public FechaPrevistaRetorno As Date
    Public IDEstadoActivo As String
    Public IDLineaMaterial As Integer

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return New String("IDLineaMaterial")
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        If oRow.Table.Columns.Contains("IDAlmacenDeposito") Then Me.IDAlmacenDeposito = oRow("IDAlmacenDeposito") & String.Empty
        If oRow.Table.Columns.Contains("IDAlmacen") Then Me.IDAlmacen = oRow("IDAlmacen") & String.Empty
        If oRow.Table.Columns.Contains("FechaPrevistaRetorno") Then Me.FechaPrevistaRetorno = Nz(oRow("FechaPrevistaRetorno"), cnMinDate)
        If oRow.Table.Columns.Contains("IDEstadoActivo") Then Me.IDEstadoActivo = Nz(oRow("IDEstadoActivo"), NegocioGeneral.ESTADOACTIVO_DISPONIBLE)
        If oRow.Table.Columns.Contains("IDLineaAlbaran") Then Me.IDLineaOrigen = oRow("IDLineaAlbaran")
        If oRow.Table.Columns.Contains("IDLineaMaterial") Then Me.IDLineaMaterial = Nz(oRow("IDLineaMaterial"), 0)
    End Sub

End Class
