Public Class PedLinCompraSolicitudCompra
    Inherits PedLinCompra

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return "IDLineaSolicitud"
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

    End Sub
End Class
