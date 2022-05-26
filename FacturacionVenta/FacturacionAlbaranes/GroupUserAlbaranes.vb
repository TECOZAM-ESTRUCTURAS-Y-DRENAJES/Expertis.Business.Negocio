'Establecer las distintas facturas a generar en función de los niveles de agrupación (Albaranes)
Public Class GroupUserAlbaranes
    Implements IGroupUser

    Public Fras(-1) As FraCabAlbaran
    Public mDteFechaFactura As Date
    Public mfvcFecha As enumfvcFechaAlbaran

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject

        Dim fralin As New FraLinAlbaran(oRow)

        Dim fraCab As FraCabAlbaran = Group
        fraCab.Add(fralin)

        If mDteFechaFactura = cnMinDate Then
            If mfvcFecha = enumfvcFechaAlbaran.fvcPrimera Then
                If fraCab.Fecha > oRow("FechaAlbaran") Then fraCab.Fecha = oRow("FechaAlbaran")
            Else
                If fraCab.Fecha < oRow("FechaAlbaran") Then fraCab.Fecha = oRow("FechaAlbaran")
            End If
        End If

    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject

        Dim fra As New FraCabAlbaran(oRow)
        If mDteFechaFactura <> cnMinDate Then fra.Fecha = mDteFechaFactura

        AddToGroupObject(oRow, fra)
        ReDim Preserve Fras(UBound(Fras) + 1)
        Fras(UBound(Fras)) = fra
        Return fra

    End Function

    Public Sub New(ByVal DteFechaFactura As Date, ByVal fvcFecha As enumfvcFechaAlbaran)
        mDteFechaFactura = DteFechaFactura
        mfvcFecha = fvcFecha
    End Sub
End Class
