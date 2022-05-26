Public Class GroupUserObraPromo
    Implements IGroupUser

    Friend fras(-1) As FraCabObraPromo

    Public mTipoFactura As enumfvcTipoFactura

    Public Sub AddToGroupObject(ByVal oRow As DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        Dim fralin As New FraLinObraPromo(oRow)
        Dim fraCab As FraCabObraPromo = Group
        fraCab.Add(fralin)
    End Sub

    Public Function NewGroupObject(ByVal oRow As DataRow) As Object Implements IGroupUser.NewGroupObject
        Dim fra As New FraCabObraPromo(oRow, mTipoFactura)
        AddToGroupObject(oRow, fra)
        ReDim Preserve fras(UBound(fras) + 1)
        fras(UBound(fras)) = fra
        Return fra
    End Function

    Public Sub New(ByVal TipoFactura As enumfvcTipoFactura)
        mTipoFactura = TipoFactura
    End Sub

End Class