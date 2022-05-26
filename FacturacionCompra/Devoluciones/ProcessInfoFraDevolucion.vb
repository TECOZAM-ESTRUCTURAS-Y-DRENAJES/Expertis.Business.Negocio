Public Class ProcessInfoFraDevolucion
    Inherits ProcessInfoFra

    Public FechaFactura As Date
    Public SuFechaFactura As Date

    Public Sub New(ByVal IDContador As String, ByVal TipoLineaDef As String, ByVal FechaFactura As Date, Optional ByVal SuFactura As String = Nothing, Optional ByVal SuFechaFactura As Date = cnMinDate)
        MyBase.New(IDContador, TipoLineaDef, SuFactura)
        Me.FechaFactura = FechaFactura
        If SuFechaFactura <> cnMinDate Then
            Me.SuFechaFactura = SuFechaFactura
        End If
    End Sub

End Class
