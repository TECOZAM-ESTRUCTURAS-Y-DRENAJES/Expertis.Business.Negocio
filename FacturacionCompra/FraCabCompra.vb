Public Class FraCabCompra
    Inherits CompraCab

    Public Dto As Double
    Public IDDireccion As Integer
    Public IDObra As Integer
    Public IDBancoPropio As String
    Public Agrupacion As enummpAgrupFactura
    Public IDTipoCompra As String
    '  Public IDProveedorBanco As Integer

    Public Sub New(ByVal oRow As IPropertyAccessor)
        MyBase.New(oRow)

        Dto = oRow("Dto")
        'Fecha = oRow("FechaAlbaran")
        IDDireccion = Nz(oRow("IDDireccion"), 0)
        If Length(oRow("IDObra")) > 0 Then IDObra = oRow("IDObra")
        If Length(oRow("IDBancoPropio")) > 0 Then IDBancoPropio = oRow("IDBancoPropio")
        Agrupacion = oRow("AgrupFactura")
        If Length(oRow("IDTipoCompra")) > 0 Then IDTipoCompra = oRow("IDTipoCompra")
        '    IDProveedorBanco = Nz(oRow("IDProveedorBanco"), 0)
    End Sub
  
End Class



