Public Class FraCab
    Inherits ComercialCab
    
    Public Dto As Double
    Public IDDireccion As Integer
    Public IDDireccionFra As Integer
    Public IDClienteBanco As Integer
    Public IDBancoPropio As String
    Public IDObra As Integer
    Public Agrupacion As enummcAgrupFactura

    Public Sub New(ByVal oRow As IPropertyAccessor)
        MyBase.New(oRow)
        IDClienteBanco = Nz(oRow("IdClienteBanco"), 0)
        IDDireccion = Nz(oRow("IDDireccion"), 0)
        IDDireccionFra = Nz(oRow("IDDireccionFra"), 0)
        If Length(oRow("IDBancoPropio")) > 0 Then IDBancoPropio = oRow("IDBancoPropio")
        If Length(oRow("IDObra")) > 0 Then IDObra = oRow("IDObra")

    End Sub
End Class

