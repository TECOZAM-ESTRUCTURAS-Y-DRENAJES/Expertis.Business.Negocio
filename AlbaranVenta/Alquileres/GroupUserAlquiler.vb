Public Class GroupUserAlquiler
    Implements IGroupUser

    Public Albs(-1) As AlbCabVentaAlquiler

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        Dim alblin As New AlbLinVentaAlquiler(oRow)

        Dim albCab As AlbCabVentaAlquiler = Group
        albCab.Add(alblin)
    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        Dim albCab As New AlbCabVentaAlquiler(oRow)

        AddToGroupObject(oRow, albCab)
        ReDim Preserve Albs(UBound(Albs) + 1)
        Albs(UBound(Albs)) = albCab
        Return albCab
    End Function

End Class
