Public Class GroupUserPlanificaciones
    Implements IGroupUser

    Public Pedidos(-1) As PedCabCompraPlanif

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        '//Creamos y añadimos una línea de pedido al grupo que representa la cabecera del pedido.
        Dim pedlin As New PedLinCompraPlanificacion(oRow)

        Dim pedCab As PedCabCompraPlanif = Group

        pedCab.Add(pedlin)
    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        '//Creamos la cabecera del pedido
        Dim Pedido As New PedCabCompraPlanif(oRow)

        AddToGroupObject(oRow, Pedido)

        ReDim Preserve Pedidos(UBound(Pedidos) + 1)
        Pedidos(UBound(Pedidos)) = Pedido

        Return Pedido
    End Function


End Class
