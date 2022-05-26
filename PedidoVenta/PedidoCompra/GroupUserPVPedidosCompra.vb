Public Class GroupUserPVPedidosCompra
    Implements IGroupUser

    Public Pedidos(-1) As PedCabVentaPedidoCompra


    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        '//Creamos la cabecera del pedido
        Dim ped As New PedCabVentaPedidoCompra(oRow)

        'AddToGroupObject(oRow, ped)

        ReDim Preserve Pedidos(Pedidos.Length)
        Pedidos(Pedidos.Length - 1) = ped

        Return ped
    End Function

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        '    '//Creamos y añadimos una línea de pedido al grupo que representa la cabecera del pedido.
        '    ' Dim pedlin As New PedLinPrograma(oRow)

        '    Dim pedCab As PedCabVentaPedidoCompra = Group
        '    pedCab.Add(oRow)
    End Sub

End Class
