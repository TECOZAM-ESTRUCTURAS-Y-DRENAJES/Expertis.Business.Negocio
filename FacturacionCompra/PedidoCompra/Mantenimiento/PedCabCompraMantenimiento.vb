Public Class PedCabCompraMantenimiento
    Inherits PedCabCompra

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        MyBase.ViewName = "vFrmMntoOTGeneraCompra"
        MyBase.Origen = enumOrigenPedidoCompra.Mnto
    End Sub

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return "IDOT"
    End Function

    Public Overrides Function FieldNOrigen() As String
        Return "NROT"
    End Function

End Class
