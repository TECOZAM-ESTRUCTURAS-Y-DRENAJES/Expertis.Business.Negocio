Public Class PedLinCompraSubcontratacion
    Inherits PedLinCompra

    Public QInterna As Double
    Public IDUDInterna As String
    Public IDUDProduccion As String
    Public FechaEntrega As Date

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
    End Sub

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return "IDOrdenRuta"
    End Function

End Class
