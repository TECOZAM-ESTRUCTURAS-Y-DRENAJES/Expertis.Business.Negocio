Public Class PedLinCompraPlanificacion
    Inherits PedLinCompra

    Public IDMarca As String
    Public IDArticulo As String
    Public IDAlmacen As String

    Public Overrides Function PrimaryKeyLinOrigen() As String
        Return String.Empty
    End Function

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        Me.IDMarca = oRow("IDMarca")
        Me.IDArticulo = oRow("IDArticulo")
        If Length(oRow("IDAlmacen")) > 0 Then Me.IDAlmacen = oRow("IDAlmacen")
    End Sub

End Class
