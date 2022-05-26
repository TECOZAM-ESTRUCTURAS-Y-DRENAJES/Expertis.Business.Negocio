Public Class ProcessInfoFraObras
    Inherits ProcessInfo

    Public TipoFacturacion As enumTipoFactura
    Public TipoAgrupacion As enummcAgrupFacturaObra
    Public CalculoSeguros As Boolean

    Public Sub New(ByVal IDContador As String, ByVal TipoFacturacion As enumTipoFactura, ByVal TipoAgrupacion As enummcAgrupFacturaObra, ByVal CalculoSeguros As Boolean)
        MyBase.New(IDContador)
        Me.TipoFacturacion = TipoFacturacion
        Me.TipoAgrupacion = TipoAgrupacion
        Me.CalculoSeguros = CalculoSeguros
    End Sub

End Class