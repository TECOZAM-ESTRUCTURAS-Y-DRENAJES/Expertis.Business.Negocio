<Serializable()> _
Public Class DataPrcCrearPedidoCompraPlanificacion

    Public Planificaciones As DataTable
    Public IDContador As String
    Public AgruparPorProveedor As Boolean

    Public Sub New(ByVal Planificaciones As DataTable, ByVal IDContador As String, ByVal AgruparPorProveedor As Boolean)
        Me.Planificaciones = Planificaciones
        Me.IDContador = IDContador
        Me.AgruparPorProveedor = AgruparPorProveedor
    End Sub

End Class
