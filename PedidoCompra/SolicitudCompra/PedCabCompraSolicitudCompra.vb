Public Class PedCabCompraSolicitudCompra
    Inherits PedCabCompra

    Public IDOperario As String

    Public Overrides Function FieldNOrigen() As String
        Return String.Empty
    End Function

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return "IDSolicitud"
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        MyBase.ViewName = "vCIGestionDeSolicitudes"
        MyBase.Origen = enumOrigenPedidoCompra.Solicitud
        If oRow.Table.Columns.Contains("IDOperario") AndAlso Length(oRow("IDOperario")) > 0 Then Me.IDOperario = oRow("IDOperario")
    End Sub

End Class
