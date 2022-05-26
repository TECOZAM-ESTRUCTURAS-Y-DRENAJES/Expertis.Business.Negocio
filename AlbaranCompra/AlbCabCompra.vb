Public Class AlbCabCompra
    Inherits CompraCab

    Public IDFormaEnvio As String
    Public IDCondicionEnvio As String
    Public IDDireccion As Integer
    Public Dto As Double
    Public IDAlmacen As String
    Public IDTipoCompra As String
    Public Fecha As Date
    Public Automatico As Boolean


    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(New DataRowPropertyAccessor(oRow))

        IDFormaEnvio = oRow("IDFormaEnvio") & String.Empty
        IDCondicionEnvio = oRow("IDCondicionEnvio") & String.Empty
        If Not oRow.IsNull("IDDireccion") Then IDDireccion = oRow("IDDireccion")
        Dto = Nz(oRow("Dto"), 0)
        IDAlmacen = oRow("IDAlmacen") & String.Empty
        IDTipoCompra = oRow("IDTipoCompra") & String.Empty
        Fecha = Today
        Automatico = True
    End Sub

End Class
   