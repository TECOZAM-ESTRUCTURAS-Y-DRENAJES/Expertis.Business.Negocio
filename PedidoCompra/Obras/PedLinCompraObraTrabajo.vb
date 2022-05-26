Public Class PedLinCompraObraTrabajo
    Inherits PedLinCompra

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return "IDTrabajo"
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
    End Sub
End Class
