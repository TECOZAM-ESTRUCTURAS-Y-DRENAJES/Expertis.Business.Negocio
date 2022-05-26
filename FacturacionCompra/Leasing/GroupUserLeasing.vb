Public Class GroupUserLeasing
    Implements IGroupUser

    Public Fras(-1) As FraCabCompraLeasing

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject

        Dim fralin As New FraLinCompraLeasing(oRow)

        Dim fraCab As FraCabCompraLeasing = Group
        fraCab.Add(fralin)

    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        Dim fra As New FraCabCompraLeasing(oRow)

        AddToGroupObject(oRow, fra)

        ReDim Preserve Fras(Fras.Length)
        Fras(Fras.Length - 1) = fra
        Return fra
    End Function
End Class
