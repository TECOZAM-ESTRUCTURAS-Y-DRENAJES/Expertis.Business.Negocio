Public Class AlbCabPedidoCompra
    Inherits AlbCabCompra

    Public IDPedido As Integer
    Public NPedido As String
    Public IDCentroCoste As String
    Public Agrupacion As enummpAgrupAlbaran

    Public Lineas(-1) As AlbLinPedidoCompra


    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        IDPedido = oRow("IDPedido")
        If oRow.Table.Columns.Contains("IDCentroCoste") AndAlso Length(oRow("IDCentroCoste")) > 0 Then IDCentroCoste = oRow("IDCentroCoste")
    End Sub

    Public Sub Add(ByVal lin As AlbLinPedidoCompra)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub

End Class
