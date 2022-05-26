
Public Class GroupUserEntregasCtaVenta
    Implements IGroupUser

    Public Fras(-1) As FraCabEntregaCta

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject

        Dim fralin As New FraLinEntregaCta(oRow)

        Dim fraCab As FraCabEntregaCta = Group
        fraCab.Add(fralin)

    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        Dim fra As New FraCabEntregaCta(New DataRowPropertyAccessor(oRow))

        AddToGroupObject(oRow, fra)

        ReDim Preserve Fras(Fras.Length)
        Fras(Fras.Length - 1) = fra
        Return fra
    End Function

End Class