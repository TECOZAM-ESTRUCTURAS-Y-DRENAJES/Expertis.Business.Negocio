<Serializable()> _
Public Class DataOfertaComercial
    Public IDLineaOfertaDetalle As Integer
    Public IDOfertaComercial As Integer
    Public PedidoCliente As String
    Public FechaEntrega As Date?

    Public Sub New(ByVal IDOfertaComercial As Integer, Optional ByVal PedidoCliente As String = Nothing, Optional ByVal FechaEntrega As Date = cnMinDate)
        Me.IDOfertaComercial = IDOfertaComercial
        If Length(PedidoCliente) > 0 Then Me.PedidoCliente = PedidoCliente
        If FechaEntrega <> cnMinDate Then Me.FechaEntrega = FechaEntrega
    End Sub

    Public Sub New(ByVal IDLineaOfertaDetalle As Integer, ByVal IDOfertaComercial As Integer, Optional ByVal PedidoCliente As String = Nothing, Optional ByVal FechaEntrega As Date = cnMinDate)
        Me.IDLineaOfertaDetalle = IDLineaOfertaDetalle
        Me.IDOfertaComercial = IDOfertaComercial
        If Length(PedidoCliente) > 0 Then Me.PedidoCliente = PedidoCliente
        If FechaEntrega <> cnMinDate Then Me.FechaEntrega = FechaEntrega
    End Sub
End Class
