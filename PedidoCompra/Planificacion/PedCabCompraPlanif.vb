Public Class PedCabCompraPlanif
    Inherits PedCabCompra

    Public DatosOrigen As DataTable

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        Me.Origen = enumOrigenPedidoCompra.Planificacion
    End Sub

    Public Overrides Function FieldNOrigen() As String
        Return String.Empty
    End Function

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return "IDProveedor"
    End Function

    Public Sub Add(ByVal lin As PedLinCompraPlanificacion)
        ReDim Preserve LineasOrigen(LineasOrigen.Length)
        LineasOrigen(LineasOrigen.Length - 1) = lin
    End Sub

End Class
