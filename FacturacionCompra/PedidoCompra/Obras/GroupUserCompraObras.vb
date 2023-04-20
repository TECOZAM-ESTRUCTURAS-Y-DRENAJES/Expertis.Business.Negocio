Public Class GroupUserCompraObras
    Implements IGroupUser

    Public Pedidos(-1) As PedCabCompraObra
    Public PorMateriales As Boolean
    Public PorTrabajos As Boolean


    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        '//Creamos y añadimos una línea de pedido al grupo que representa la cabecera del pedido.
        Dim pedlin As PedLinCompra
        If Me.PorMateriales Then
            pedlin = New PedLinCompraObraMaterial(oRow)
        ElseIf Me.PorTrabajos Then
            pedlin = New PedLinCompraObraTrabajo(oRow)
        End If

        Dim pedCab As PedCabCompraObra = Group
        pedCab.Add(pedlin)
    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        '//Creamos la cabecera del pedido
        Dim Pedido As New PedCabCompraObra(oRow, Me.PorMateriales, Me.PorTrabajos)

        AddToGroupObject(oRow, Pedido)

        ReDim Preserve Pedidos(UBound(Pedidos) + 1)
        Pedidos(UBound(Pedidos)) = Pedido

        Return Pedido
    End Function

    Public Sub New(ByVal PorMateriales As Boolean, ByVal PorTrabajos As Boolean)
        Me.PorMateriales = PorMateriales
        Me.PorTrabajos = PorTrabajos
    End Sub

End Class
