<Serializable()> _
Public Class DataPrcFacturacionObras

    Public IDVencimiento() As Object
    Public IDContador As String
    Public TipoAgrupacion As enummcAgrupFacturaObra
    Public TipoFacturacion As enumTipoFactura
    Public CalculoSeguros As Boolean = False

    Public Sub New(ByVal IDVencimiento() As Object, ByVal IDContador As String, ByVal TipoAgrupacion As enummcAgrupFacturaObra, ByVal TipoFacturacion As enumTipoFactura)
        Me.IDVencimiento = IDVencimiento
        Me.IDContador = IDContador
        Me.TipoAgrupacion = TipoAgrupacion
        Me.TipoFacturacion = TipoFacturacion
    End Sub
    Public Sub New(ByVal IDVencimiento() As Object, ByVal IDContador As String, ByVal TipoAgrupacion As enummcAgrupFacturaObra, ByVal TipoFacturacion As enumTipoFactura, ByVal CalculoSeguros As Boolean)
        Me.IDVencimiento = IDVencimiento
        Me.IDContador = IDContador
        Me.TipoAgrupacion = TipoAgrupacion
        Me.TipoFacturacion = TipoFacturacion
        Me.CalculoSeguros = CalculoSeguros
    End Sub
End Class