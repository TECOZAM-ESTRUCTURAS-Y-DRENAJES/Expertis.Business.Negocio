Public Class GroupUserObraVencimiento
    Implements IGroupUser

    Public Fras(-1) As FraCabObra
    Public mIDCentroGestion As String
    Public mIDCondicionPago As String

    Public Sub AddToGroupObject(ByVal oRow As System.Data.DataRow, ByVal Group As Object) Implements IGroupUser.AddToGroupObject

        Dim fralin As New FraLinVencimiento(oRow)

        Dim fraCab As FraCabVencimiento = Group
        fraCab.Add(fralin)

    End Sub

    Public Function NewGroupObject(ByVal oRow As System.Data.DataRow) As Object Implements IGroupUser.NewGroupObject

        Dim fra As New FraCabVencimiento(oRow)
        If Len(fra.IDCentroGestion) = 0 Then fra.IDCentroGestion = mIDCentroGestion
        If Len(fra.IDCondicionPago) = 0 Then fra.IDCondicionPago = mIDCondicionPago

        AddToGroupObject(oRow, fra)
        ReDim Preserve Fras(Fras.Length)
        Fras(Fras.Length - 1) = fra
        Return fra

    End Function

    Public Sub New(ByVal IDCentroGestion As String, ByVal IDCondicionPago As String)
        Me.mIDCentroGestion = IDCentroGestion
        Me.mIDCondicionPago = IDCondicionPago
    End Sub
End Class
