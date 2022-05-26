<Serializable()> _
Public Class DataPrcActualizarFactura

    Public TipoFactura As enumTipoFactura
    Public RstFacturacion As ResultFacturacion

    Public Sub New(ByVal RstFacturacion As ResultFacturacion, Optional ByVal TipoFactura As enumTipoFactura = enumTipoFactura.tfNormal)
        Me.RstFacturacion = RstFacturacion
        Me.TipoFactura = TipoFactura
    End Sub

End Class
