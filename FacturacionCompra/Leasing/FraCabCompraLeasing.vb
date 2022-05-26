Public Class FraCabCompraLeasing
    Inherits FraCabCompra

    'Public IDFactura As Integer
    'Public NFactura As String
    Public SuFactura As String
    Public RazonSocial As String
    Public CIFProveedor As String
    Public IDContador As String
    Public IDDiaPago As String
    Public IDProveedorBanco As String
    'Public FechaFactura As Date

    'Public NRegistro As Integer
    'Public IdPago As Integer

    Public Lineas(-1) As FraLinCompraLeasing

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(New DataRowPropertyAccessor(oRow))
        Fecha = Nz(oRow("FechaVencimiento"), Today)
        RazonSocial = oRow("RazonSocial") & String.Empty
        CIFProveedor = oRow("CIFProveedor") & String.Empty
        RazonSocial = oRow("RazonSocial") & String.Empty
        IDContador = oRow("IDContadorCargo") & String.Empty
        IDDiaPago = oRow("IDDiaPago") & String.Empty


        'IDAlbaran = oRow("IDAlbaran")
        'NAlbaran = oRow("NAlbaran")
        ''Fecha = oRow("FechaAlbaran")
        'Texto = oRow("Texto") & String.Empty
    End Sub

    Public Sub Add(ByVal lin As FraLinCompraLeasing)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub

End Class
