Public Class FraCabCompraNuevoGasto
    Inherits FraCabCompra

    Public RazonSocial As String
    Public CIF As String
    Public IDDiaPago As String
    Public IDTipoAsiento As Integer
    Public lineas As New List(Of FraLinCompraNuevoGasto)

    Public Sub New(ByVal oRow As IPropertyAccessor)
        MyBase.New(oRow)
        Fecha = oRow("FechaFactura")
        RazonSocial = oRow("RazonSocial") & String.Empty
        CIF = oRow("CIF") & String.Empty
        IDDiaPago = oRow("IDDiaPago") & String.Empty
        IDTipoAsiento = oRow("IDTipoAsiento")
    End Sub

End Class
