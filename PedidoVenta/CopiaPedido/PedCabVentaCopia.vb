Public Class PedCabVentaCopia
    Inherits PedCab

    Public IDPedido As Integer

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        IDPedido = oRow("IDPedido")
        Me.Origen = enumOrigenPedido.Copia
    End Sub

End Class
