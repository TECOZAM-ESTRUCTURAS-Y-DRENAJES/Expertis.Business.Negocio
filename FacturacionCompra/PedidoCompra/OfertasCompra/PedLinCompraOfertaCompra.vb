Public Class PedLinCompraOfertaCompra
    Inherits PedLinCompra

    Public IDUDCompra As String
    Public QInterna As Double

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return "IDLineaOferta"
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        Me.IDUDCompra = oRow("IDUDCompra") & String.Empty
        Me.QInterna = Nz(oRow("QOferta"), 0)
    End Sub
End Class
