Public Class PedLinCompraProgramaCompra
    Inherits PedLinCompra

    Public FechaConfirmacion As Date?

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return "IDLineaPrograma"
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
    End Sub

End Class
