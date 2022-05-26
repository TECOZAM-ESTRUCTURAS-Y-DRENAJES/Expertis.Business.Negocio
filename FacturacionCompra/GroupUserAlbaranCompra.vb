Public Class GroupUserAlbaranCompra
    Implements IGroupUser

    Public Fras(-1) As FraCabCompra
    Public mDteFechaFactura As Date

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject

        Dim fralin As New FraLinCompraAlbaran(oRow)

        Dim fraCab As FraCabCompraAlbaran = Group
        fraCab.Add(fralin)

        If mDteFechaFactura = cnMinDate Then
            fraCab.Fecha = Today 'oRow("FechaAlbaran")
        End If

    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject

        Dim fra As New FraCabCompraAlbaran(oRow)
        'If mDteFechaFactura <> cnMinDate Then fra.Fecha = mDteFechaFactura

        AddToGroupObject(oRow, fra)
        ReDim Preserve Fras(UBound(Fras) + 1)
        Fras(UBound(Fras)) = fra
        Return fra

    End Function

    Public Sub New() '(ByVal DteFechaFactura As Date)
        ' mDteFechaFactura = DteFechaFactura
    End Sub
End Class
