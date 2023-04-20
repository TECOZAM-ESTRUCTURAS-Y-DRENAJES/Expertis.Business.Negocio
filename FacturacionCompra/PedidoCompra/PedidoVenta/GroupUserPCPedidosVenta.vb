Public Class GroupUserPCPedidosVenta
    Implements IGroupUser

    Public Pedidos(-1) As PedCabCompraPedidoVenta

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        Dim pedLin As New PedLinCompraPedidoVenta(oRow)
        Dim pedcab As PedCabCompraPedidoVenta = Group
        pedcab.Add(pedLin)
    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        Dim ped As New PedCabCompraPedidoVenta(oRow)
        AddToGroupObject(oRow, ped)
        ReDim Preserve Pedidos(Pedidos.Length)
        Pedidos(Pedidos.Length - 1) = ped
        Return ped
    End Function

End Class

