Public Class PedCabCompraOfertaComercial
    Inherits PedCabCompra

    Public PedidoCliente As String


    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        If Length(oRow("IDProveedor")) > 0 Then Me.IDProveedor = oRow("IDProveedor")
        If Length(Me.IDProveedor) = 0 AndAlso Length(oRow("IDEmpresa")) > 0 Then Me.IDProveedor = oRow("IDEmpresa")

        Fecha = Today
        MyBase.ViewName = "vfrmOfertaComercialTratamientoCompra"
        MyBase.Origen = enumOrigenPedidoCompra.OfertaComercial
    End Sub

    Public Overrides Function FieldNOrigen() As String
        Return New String("NumOferta")
    End Function

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return New String("IDOfertaComercial")
    End Function

End Class
