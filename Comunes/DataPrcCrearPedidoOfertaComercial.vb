<Serializable()> _
Public Class DataPrcCrearPedidoOfertaComercial

    Public Ofertas() As DataOfertaComercial
    Public IDContador As String
    Public Detalle As Boolean

    Public Sub New(ByVal Ofertas() As DataOfertaComercial, ByVal IDContador As String, Optional ByVal Detalle As Boolean = False)
        Me.Ofertas = Ofertas
        Me.IDContador = IDContador
        Me.Detalle = Detalle
    End Sub

End Class

