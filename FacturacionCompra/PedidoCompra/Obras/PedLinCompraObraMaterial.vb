Public Class PedLinCompraObraMaterial
    Inherits PedLinCompra

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return "IDLineaMaterial"
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
    End Sub
End Class
