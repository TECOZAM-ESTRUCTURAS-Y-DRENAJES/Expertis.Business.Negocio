Public Class PedLinCompraPedidoVenta
    Inherits PedLinCompra

    Public Cantidad2 As Double?

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return "IDLineaPedido"
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
    End Sub

End Class
