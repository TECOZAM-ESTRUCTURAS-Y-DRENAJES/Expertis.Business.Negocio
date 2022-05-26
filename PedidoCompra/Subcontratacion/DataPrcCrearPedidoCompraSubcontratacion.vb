<Serializable()> _
Public Class DataPrcCrearPedidoCompraSubcontratacion

    Public Subcontrataciones() As DataSubcontratacion
    Public IDContador As String
    Public AgruparPorProveedor As Boolean

    Public Sub New(ByVal Subcontrataciones() As DataSubcontratacion, ByVal IDContador As String, ByVal AgruparPorProveedor As Boolean)
        Me.Subcontrataciones = Subcontrataciones
        Me.IDContador = IDContador
        Me.AgruparPorProveedor = AgruparPorProveedor
    End Sub

End Class
