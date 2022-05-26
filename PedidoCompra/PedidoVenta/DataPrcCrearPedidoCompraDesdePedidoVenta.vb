Imports System.Collections.Generic

<Serializable()> _
Public Class DataPrcCrearPedidoCompraDesdePedidoVenta

    Public IDContador As String
    Public Propuestas As DataTable

    Public Sub New(ByVal Propuestas As DataTable, ByVal IDContador As String)
        Me.Propuestas = Propuestas
        Me.IDContador = IDContador
    End Sub

End Class

'//provisional: esta clase en un fichero diferente (proyecto esta cogido por otro usuario)
<Serializable()> _
Public Class DataPrcPropuestaPedidoCompraDesdePedidoVenta
    Implements IEquatable(Of DataPrcPropuestaPedidoCompraDesdePedidoVenta)

    Public IDPedido As Integer
    Public Lineas As List(Of DataLineaPrcPropuestaPedidoCompraDesdePedidoVenta)
    Public PedidoCompleto As Boolean

    Public Sub New()
        Me.Lineas = New List(Of DataLineaPrcPropuestaPedidoCompraDesdePedidoVenta)
        PedidoCompleto = True
    End Sub

    Public Function Equals(ByVal other As DataPrcPropuestaPedidoCompraDesdePedidoVenta) As Boolean Implements System.IEquatable(Of DataPrcPropuestaPedidoCompraDesdePedidoVenta).Equals
        Return IDPedido.Equals(other.IDPedido)
    End Function
End Class

<Serializable()> _
Public Class DataLineaPrcPropuestaPedidoCompraDesdePedidoVenta
    Public IDLineaPedido As Integer
    Public QPedida As Double
End Class

