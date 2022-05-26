Public Class GroupUserPedidos
    Implements IGroupUser

    Public Albs(-1) As AlbCabVentaPedido

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        Dim alblin As New AlbLinVentaPedido(oRow)

        Dim albCab As AlbCabVentaPedido = Group
        albCab.Add(alblin)

    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject

        Dim albCab As New AlbCabVentaPedido(oRow)

        AddToGroupObject(oRow, albCab)
        ReDim Preserve Albs(UBound(Albs) + 1)
        Albs(UBound(Albs)) = albCab
        Return albCab

    End Function

End Class
