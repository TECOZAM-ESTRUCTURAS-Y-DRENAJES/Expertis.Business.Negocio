Public Class GroupUserSubcontrataciones
    Implements IGroupUser

    Public Pedidos(-1) As PedCabCompraSubcontratacion
    Public mTipoCompra As String

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        '//Creamos y añadimos una línea de pedido al grupo que representa la cabecera del pedido.
        Dim pedlin As New PedLinCompraSubcontratacion(oRow)

        Dim pedCab As PedCabCompraSubcontratacion = Group
        pedCab.Add(pedlin)
    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        '//Creamos la cabecera del pedido
        Dim Pedido As New PedCabCompraSubcontratacion(oRow)

        AddToGroupObject(oRow, Pedido)

        ReDim Preserve Pedidos(UBound(Pedidos) + 1)
        Pedidos(UBound(Pedidos)) = Pedido

        Return Pedido
    End Function

    Public Sub New(ByVal TipoCompra As string)
        Me.mTipoCompra = TipoCompra
    End Sub
End Class
