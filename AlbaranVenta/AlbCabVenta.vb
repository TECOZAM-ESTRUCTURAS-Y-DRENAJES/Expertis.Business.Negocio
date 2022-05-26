Public MustInherit Class AlbCabVenta
    Inherits ComercialCab

    Public MustOverride Function PrimaryKeyCabOrigen() As String
    Public MustOverride Function FieldNOrigen() As String

    Public IDOrigen As Integer
    Public NOrigen As String

    Public Origen As enumOrigenAlbaranVenta
    Public LineasOrigen(-1) As AlbLinVenta

    Public ViewName As String

    Public IDFormaEnvio As String
    Public IDCondicionEnvio As String
    Public IdDireccion As Integer?
    Public IDDireccionFra As Integer?
    Public IDBancoPropio As String
    Public Dto As Double
    Public IDAlmacen As String
    Public IDDAA As Guid?
    Public NDAA As String
    Public IDDAABaseDatos As Guid?
    Public AadReferenceCode As String
    ' Public IDObra As Integer?

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(New DataRowPropertyAccessor(oRow))

        IDOrigen = oRow(PrimaryKeyCabOrigen)
        If Length(FieldNOrigen) > 0 AndAlso Length(oRow(FieldNOrigen)) > 0 Then NOrigen = oRow(FieldNOrigen)

        If oRow.Table.Columns.Contains("IDAlmacen") AndAlso Length(oRow("IDAlmacen")) > 0 Then IDAlmacen = oRow("IDAlmacen")
        If oRow.Table.Columns.Contains("IDDAA") AndAlso Not IsDBNull(oRow("IDDAA")) AndAlso Not CType(oRow("IDDAA"), Guid).Equals(Guid.Empty) Then
            IDDAA = oRow("IDDAA")
            If oRow.Table.Columns.Contains("NDAA") Then
                NDAA = oRow("NDAA")
            End If
            If oRow.Table.Columns.Contains("IDDAABaseDatos") AndAlso Not IsDBNull(oRow("IDDAABaseDatos")) AndAlso Not CType(oRow("IDDAABaseDatos"), Guid).Equals(Guid.Empty) Then
                IDDAABaseDatos = oRow("IDDAABaseDatos")
            End If
            If oRow.Table.Columns.Contains("AadReferenceCode") AndAlso Length(oRow("AadReferenceCode")) > 0 Then
                AadReferenceCode = oRow("AadReferenceCode")
            End If
        End If

        'If Length(oRow("IDObra")) > 0 Then IDObra = oRow("IDObra")
        'IDClienteBanco = Nz(oRow("IdClienteBanco"), 0)
    End Sub

    Public Sub Add(ByVal lin As AlbLinVenta)
        ReDim Preserve LineasOrigen(LineasOrigen.Length)
        LineasOrigen(LineasOrigen.Length - 1) = lin
    End Sub

End Class
 