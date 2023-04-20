<serializable()> _
Public Class DataPrcCrearPedidoCompraSolicitudCompra

    Public Solicitudes() As DataSolicitudCompra
    Public IDContador As String

    Public Sub New(ByVal Solicitudes() As DataSolicitudCompra, ByVal IDContador As String)
        Me.Solicitudes = Solicitudes
        Me.IDContador = IDContador
    End Sub

End Class