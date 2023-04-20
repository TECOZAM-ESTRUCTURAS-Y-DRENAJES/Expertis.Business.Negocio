Public Class GroupUserProgramaCompra
    Implements IGroupUser

    Public Peds(-1) As PedCabCompraProgramaCompra


    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        '//Creamos la cabecera del pedido
        Dim ped As New PedCabCompraProgramaCompra(oRow)

        AddToGroupObject(oRow, ped)

        ReDim Preserve Peds(UBound(Peds) + 1)
        Peds(UBound(Peds)) = ped

        Return ped
    End Function

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        '//Creamos y añadimos una línea de pedido al grupo que representa la cabecera del pedido.
        Dim pedlin As New PedLinCompraProgramaCompra(oRow)

        Dim pedCab As PedCabCompraProgramaCompra = Group
        pedCab.Add(pedlin)
    End Sub

End Class

