Public Class PedCab
    Inherits ComercialCab

    Public Origen As enumOrigenPedido
    Public Agrupacion As enummcAgrupPedido
    Public PedidoCliente As String

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(New DataRowPropertyAccessor(oRow))

        If oRow.Table.Columns.Contains("AgrupPedido") Then
            Agrupacion = oRow("AgrupPedido")
        Else
            Agrupacion = enummcAgrupPedido.mcCliente
        End If
    End Sub

End Class
