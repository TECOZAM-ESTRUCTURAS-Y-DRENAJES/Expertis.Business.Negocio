Public Class GroupUserCompraOfertasComerciales
    Implements IGroupUser


    Public Peds(-1) As PedCabCompraOfertaComercial

    'PrimaryKeyLinOrigen

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        '//Creamos la cabecera del pedido
        Dim ped As New PedCabCompraOfertaComercial(oRow)

        AddToGroupObject(oRow, ped)

        ReDim Preserve Peds(Peds.Length)
        Peds(Peds.Length - 1) = ped

        Return ped
    End Function

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        '//Creamos y añadimos una línea de pedido al grupo que representa la cabecera del pedido.
        Dim pedlin As New PedLinCompraOfertaComercial(oRow)

        Dim pedCab As PedCabCompraOfertaComercial = Group
        pedCab.Add(pedlin)
    End Sub

End Class
