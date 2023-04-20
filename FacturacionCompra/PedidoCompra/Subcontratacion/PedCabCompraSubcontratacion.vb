Public Class PedCabCompraSubcontratacion
    Inherits PedCabCompra

    Public IDAlmacen As String

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return "IDOrden"
    End Function

    Public Overrides Function FieldNOrigen() As String
        Return "NOrden"
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        MyBase.ViewName = "vCTLCIEnvioASubcontratacion"
        MyBase.Origen = enumOrigenPedidoCompra.Subcontratacion
        If Length(oRow("IDAlmacen")) > 0 Then Me.IDAlmacen = oRow("IDAlmacen")
    End Sub

End Class