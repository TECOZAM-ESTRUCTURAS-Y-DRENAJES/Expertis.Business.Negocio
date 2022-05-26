Public Class GroupUserObraCertificacion
    Implements IGroupUser

    Public fras(-1) As FraCabObraCertificacion

    Public mIDCentroGestion As String
    Public Sub AddToGroupObject(ByVal oRow As DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject
        Dim fralin As New FraLinObraCertificacion(oRow)
        Dim fraCab As FraCabObraCertificacion = Group
        fraCab.Add(fralin)
    End Sub

    Public Function NewGroupObject(ByVal oRow As DataRow) As Object Implements IGroupUser.NewGroupObject
        Dim fra As New FraCabObraCertificacion(New DataRowPropertyAccessor(oRow))
        AddToGroupObject(oRow, fra)
        ReDim Preserve fras(UBound(fras) + 1)
        fras(UBound(fras)) = fra
        Return fra
    End Function

    Public Sub New(ByVal IDCentroGestion As String)
        Me.mIDCentroGestion = IDCentroGestion
    End Sub

End Class