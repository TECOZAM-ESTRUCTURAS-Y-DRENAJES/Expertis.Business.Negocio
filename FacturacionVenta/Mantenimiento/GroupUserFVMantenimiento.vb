Public Class GroupUserFVMantenimiento
    Implements IGroupUser

    Public Fras(-1) As FraCabMnto

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject

        Dim fralin As New FraLinMnto(oRow)

        Dim fraCab As FraCabMnto = Group
        fraCab.Add(fralin)

    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject
        Dim fra As New FraCabMnto(New DataRowPropertyAccessor(oRow))

        AddToGroupObject(oRow, fra)

        ReDim Preserve Fras(Fras.Length)
        Fras(Fras.Length - 1) = fra

        Return fra
    End Function


End Class
