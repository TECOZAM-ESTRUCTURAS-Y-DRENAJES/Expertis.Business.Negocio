Public Class AlbLinVentaObras
    Inherits AlbLinVenta

    Public IDCentroGestion As String
    Public IDTipoIVA As String

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return New String("IDLineaMaterial")
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        If oRow.Table.Columns.Contains("IDCentroGestion") Then Me.IDCentroGestion = oRow("IDCentroGestion") & String.Empty
        If oRow.Table.Columns.Contains("IDTipoIVA") Then Me.IDTipoIVA = oRow("IDTipoIVA") & String.Empty
    End Sub

End Class
