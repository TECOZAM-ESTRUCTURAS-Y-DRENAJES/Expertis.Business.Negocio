Public Class FraCabCompraEntregaCta
    Inherits FraCabCompra

    Public IDEntrega As Integer
    Public CContableProveedor As String
    Public IDTipoPago As String
    Public TipoEntrega As Integer

    Public Lineas(-1) As FraLinEntregaCta

    Public Sub New(ByVal oRow As IPropertyAccessor)
        MyBase.New(oRow)
        Me.IDEntrega = oRow("IDEntrega")
        Me.Fecha = oRow("FechaEntrega")
        Me.CContableProveedor = oRow("CCClienteProveedor")
        Me.IDTipoPago = oRow("IDTipoCobroPago")
        Me.TipoEntrega = oRow("TipoEntrega")
        If Length(oRow("IDObra")) > 0 Then Me.IDObra = oRow("IDObra")
    End Sub

    Public Sub Add(ByVal lin As FraLinEntregaCta)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub
End Class
