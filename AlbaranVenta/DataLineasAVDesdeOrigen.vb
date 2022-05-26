Public Class DataLineasAVDesdeOrigen
    Public Row As DataRow
    Public Origen As DataRow
    Public Cantidad As Double
    'Public CantidadUd As Double
    'Public IDUDInterna As String
    'Public IDUDMedida As String
    Public NSerie As String
    Public IDEstadoActivo As String
    Public IDOperario As String
    Public Ubicacion As String

    Public Doc As DocumentoAlbaranVenta
    Public AlbLin As AlbLinVenta

    Public Sub New(ByVal Row As DataRow, ByVal Origen As DataRow, ByVal Doc As DocumentoAlbaranVenta, ByVal AlbLin As AlbLinVenta, ByVal Cantidad As Double, Optional ByVal RowSerie As DataRow = Nothing)
        Me.Row = Row
        Me.Origen = Origen
        Me.Doc = Doc
        Me.AlbLin = AlbLin
        Me.Cantidad = Cantidad
        If Not RowSerie Is Nothing Then
            Me.NSerie = RowSerie("NSerie")
            Me.Cantidad = IIf(AlbLin.QaServir > 0, 1, -1)
            Me.IDEstadoActivo = RowSerie("IDEstadoActivo")
            Me.IDOperario = RowSerie("IDOperario") & String.Empty
            If RowSerie.Table.Columns.Contains("Ubicacion") AndAlso Length(RowSerie("Ubicacion")) Then Me.Ubicacion = RowSerie("Ubicacion")
        End If
    End Sub
End Class
